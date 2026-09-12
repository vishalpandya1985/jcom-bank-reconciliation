import os
import shutil
import uuid
import json
from datetime import datetime, timezone

from fastapi import FastAPI, Request, UploadFile, File, Form, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

import database as db
import auth
from report_builder import run_reconciliation, build_workbook_from_detail
from export_formats import export_csv, export_xlsx, export_pdf
from reconcile_engine import load_hdfc, find_candidates, rebuild_gl_row_from_entry

BASE_DIR = os.path.dirname(__file__)
UPLOADS_DIR = os.path.join(BASE_DIR, 'uploads')  # transient - cleared after every run, fine to lose on restart
REPORTS_DIR = os.environ.get('REPORTS_DIR', os.path.join(BASE_DIR, 'reports'))  # persistent - past reports live here
os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

app = FastAPI(title='JCOM Bank Reconciliation')
app.mount('/static', StaticFiles(directory=os.path.join(BASE_DIR, 'static')), name='static')
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, 'templates'))
templates.env.filters['from_json'] = json.loads

db.init_db()


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def get_current_user(request: Request):
    token = request.cookies.get(auth.SESSION_COOKIE_NAME)
    user_id = auth.read_session_token(token)
    if not user_id:
        return None
    return db.get_user_by_id(user_id)


def require_login(request: Request):
    user = get_current_user(request)
    if not user:
        return None
    return user


def require_login_password_ok(request: Request):
    """Like require_login, but also redirects to the forced change-password
    page if the account still has a temporary/admin-set password. Returns
    (user, None) normally, or (None, RedirectResponse) when the caller
    should return that redirect instead of rendering the page."""
    user = require_login(request)
    if not user:
        return None, RedirectResponse('/login', status_code=302)
    if user['must_change_password']:
        return None, RedirectResponse('/change-password?forced=1', status_code=302)
    return user, None


# ---------------------------------------------------------------------------
# Auth routes
# ---------------------------------------------------------------------------

@app.get('/', response_class=HTMLResponse)
def root(request: Request):
    user = get_current_user(request)
    if user:
        return RedirectResponse('/dashboard', status_code=302)
    if not db.any_users_exist():
        return RedirectResponse('/signup', status_code=302)
    return RedirectResponse('/login', status_code=302)


@app.get('/signup', response_class=HTMLResponse)
def signup_form(request: Request):
    # Only allowed to bootstrap the very first (admin) account. After that,
    # new team members are added by an admin from the Users page.
    if db.any_users_exist():
        return RedirectResponse('/login', status_code=302)
    return templates.TemplateResponse(request, 'signup.html', {})


@app.post('/signup')
def signup_submit(request: Request, username: str = Form(...), password: str = Form(...),
                   display_name: str = Form('')):
    if db.any_users_exist():
        raise HTTPException(403, 'Signup is closed. Ask your admin to create your account.')
    # the first account sets its own real password, so it doesn't need to be
    # forced through a change-password step the way admin-created accounts do
    db.create_user(username.strip().lower(), auth.hash_password(password), display_name or username,
                    is_admin=1, must_change_password=0)
    user = db.get_user_by_username(username.strip().lower())
    resp = RedirectResponse('/dashboard', status_code=302)
    token = auth.create_session_token(user['id'])
    resp.set_cookie(auth.SESSION_COOKIE_NAME, token, httponly=True, max_age=auth.SESSION_MAX_AGE)
    return resp


@app.get('/login', response_class=HTMLResponse)
def login_form(request: Request, error: str = ''):
    return templates.TemplateResponse(request, 'login.html', {'error': error})


MAX_LOGIN_ATTEMPTS = 5


@app.post('/login')
def login_submit(username: str = Form(...), password: str = Form(...)):
    user = db.get_user_by_username(username.strip().lower())
    if not user:
        return RedirectResponse('/login?error=Invalid+username+or+password', status_code=302)

    if user['is_locked']:
        return RedirectResponse(
            '/login?error=This+account+is+locked+after+too+many+failed+attempts.+Ask+your+admin+to+unlock+it.',
            status_code=302)

    if not auth.verify_password(password, user['password_hash']):
        attempts, now_locked = db.record_failed_login(user['id'], MAX_LOGIN_ATTEMPTS)
        if now_locked:
            return RedirectResponse(
                '/login?error=Too+many+failed+attempts.+This+account+is+now+locked+-+ask+your+admin+to+unlock+it.',
                status_code=302)
        remaining = MAX_LOGIN_ATTEMPTS - attempts
        return RedirectResponse(
            f'/login?error=Invalid+username+or+password+({remaining}+attempt{"s" if remaining != 1 else ""}+remaining)',
            status_code=302)

    db.reset_failed_attempts(user['id'])
    resp = RedirectResponse('/change-password?forced=1' if user['must_change_password'] else '/dashboard',
                             status_code=302)
    token = auth.create_session_token(user['id'])
    resp.set_cookie(auth.SESSION_COOKIE_NAME, token, httponly=True, max_age=auth.SESSION_MAX_AGE)
    return resp


@app.get('/logout')
def logout():
    resp = RedirectResponse('/login', status_code=302)
    resp.delete_cookie(auth.SESSION_COOKIE_NAME)
    return resp


@app.get('/forgot-password', response_class=HTMLResponse)
def forgot_password(request: Request):
    return templates.TemplateResponse(request, 'forgot_password.html', {})


@app.get('/change-password', response_class=HTMLResponse)
def change_password_form(request: Request, forced: str = '', error: str = ''):
    user = require_login(request)
    if not user:
        return RedirectResponse('/login', status_code=302)
    return templates.TemplateResponse(request, 'change_password.html', {
        'user': user, 'forced': bool(forced) and bool(user['must_change_password']), 'error': error,
    })


@app.post('/change-password')
def change_password_submit(request: Request, new_password: str = Form(...),
                            confirm_password: str = Form(...), current_password: str = Form('')):
    user = require_login(request)
    if not user:
        return RedirectResponse('/login', status_code=302)

    if new_password != confirm_password:
        return RedirectResponse('/change-password?error=New+passwords+do+not+match', status_code=302)
    if len(new_password) < 6:
        return RedirectResponse('/change-password?error=Password+must+be+at+least+6+characters', status_code=302)

    # the forced first-change flow (right after login) skips re-entering the
    # password the person just typed; a voluntary change later always
    # requires the current password to prevent a hijacked session from
    # silently taking over the account
    if not user['must_change_password']:
        if not current_password or not auth.verify_password(current_password, user['password_hash']):
            return RedirectResponse('/change-password?error=Current+password+is+incorrect', status_code=302)

    db.set_password(user['id'], auth.hash_password(new_password), must_change_password=0)
    return RedirectResponse('/dashboard', status_code=302)


# ---------------------------------------------------------------------------
# Dashboard / upload
# ---------------------------------------------------------------------------

@app.get('/dashboard', response_class=HTMLResponse)
def dashboard(request: Request):
    user, redirect = require_login_password_ok(request)
    if redirect:
        return redirect
    return templates.TemplateResponse(request, 'dashboard.html', {'user': user, 'active': 'dashboard'})


@app.post('/api/reconcile')
async def api_reconcile(
    request: Request,
    run_label: str = Form(''),
    file_3493: UploadFile = File(...),
    file_3496: UploadFile = File(...),
    file_345051: UploadFile = File(...),
    file_hdfc: UploadFile = File(...),
):
    user = require_login(request)
    if not user:
        raise HTTPException(401, 'Not logged in')

    run_uuid = uuid.uuid4().hex[:12]
    run_upload_dir = os.path.join(UPLOADS_DIR, run_uuid)
    os.makedirs(run_upload_dir, exist_ok=True)

    incoming = {
        '3493': file_3493, '3496': file_3496, '345051': file_345051, 'hdfc': file_hdfc,
    }
    saved_paths = {}
    original_names = {}
    try:
        for key, upload in incoming.items():
            dest = os.path.join(run_upload_dir, f'{key}_{upload.filename}')
            with open(dest, 'wb') as f:
                shutil.copyfileobj(upload.file, f)
            saved_paths[key] = dest
            original_names[key] = upload.filename

        report_filename = f'{run_uuid}_reconciliation_report.xlsx'
        report_path = os.path.join(REPORTS_DIR, report_filename)
        summary, detail = run_reconciliation(saved_paths, report_path)

        detail_filename = f'{run_uuid}_details.json'
        detail_path = os.path.join(REPORTS_DIR, detail_filename)
        with open(detail_path, 'w') as f:
            json.dump(detail, f)

        label = run_label.strip() or datetime.now(timezone.utc).strftime('Run %Y-%m-%d %H:%M UTC')
        run_id = db.save_run(user['id'], label, summary, report_path, original_names, detail_path)

        return JSONResponse({'ok': True, 'run_id': run_id, 'summary': summary})
    except Exception as e:
        return JSONResponse({'ok': False, 'error': str(e)}, status_code=400)
    finally:
        # uploaded source CSVs aren't needed after the report is built
        shutil.rmtree(run_upload_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# History
# ---------------------------------------------------------------------------

@app.get('/history', response_class=HTMLResponse)
def history_page(request: Request):
    user, redirect = require_login_password_ok(request)
    if redirect:
        return redirect
    runs = db.list_runs(user_id=None if user['is_admin'] else user['id'])
    return templates.TemplateResponse(request, 'history.html', {'user': user, 'runs': runs, 'active': 'history'})


@app.get('/reports/{run_id}')
def download_report(request: Request, run_id: int):
    user = require_login(request)
    if not user:
        return RedirectResponse('/login', status_code=302)
    run = db.get_run(run_id)
    if not run:
        raise HTTPException(404, 'Run not found')
    if not user['is_admin'] and run['user_id'] != user['id']:
        raise HTTPException(403, 'Not your report')
    if not os.path.exists(run['report_path']):
        raise HTTPException(404, 'Report file missing on server')
    filename = f"reconciliation_{run['created_at'][:10]}.xlsx"
    return FileResponse(run['report_path'], filename=filename,
                         media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


def _load_run_detail_or_403(request: Request, run_id: int):
    user = require_login(request)
    if not user:
        return None, None, RedirectResponse('/login', status_code=302)
    run = db.get_run(run_id)
    if not run:
        raise HTTPException(404, 'Run not found')
    if not user['is_admin'] and run['user_id'] != user['id']:
        raise HTTPException(403, 'Not your report')
    if not run['detail_path'] or not os.path.exists(run['detail_path']):
        raise HTTPException(404, 'Entry-level detail is not available for this run (it may predate this feature) - only the Excel report can be downloaded.')
    with open(run['detail_path']) as f:
        detail = json.load(f)
    return user, run, detail


FLAG_KEYWORDS = [
    ('duplicate group', 'DUP'),
    ('cleared next day', 'NEXTDAY'),
    ('manually reviewed', 'MANUAL'),
    ('reference no', 'REF'),
    ('embedded transaction code', 'CODE'),
    ('batch label', 'BATCH'),
    ('beneficiary/remitter name', 'NAME'),
    ('multiple', 'MULTI'),
    ('missing', 'MISS'),
    ('internal gl sweep', 'N/A'),
]

WEEKDAY_NAMES = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']


def _flag_for(note, status):
    text = (note or '').lower() + ' ' + (status or '').lower()
    for keyword, tag in FLAG_KEYWORDS:
        if keyword in text:
            return tag
    return '-'


def _weekday_for(date_str):
    # dates come as dd-mm-yyyy from the GL export
    for fmt in ('%d-%m-%Y', '%d/%m/%Y'):
        try:
            return WEEKDAY_NAMES[datetime.strptime(date_str, fmt).weekday()]
        except (ValueError, TypeError):
            continue
    return '-'


def _enrich_entries(entries):
    out = []
    for i, e in enumerate(entries, start=1):
        enriched = dict(e)
        enriched['seq'] = i
        enriched['flag'] = _flag_for(e.get('note'), e.get('status'))
        enriched['day'] = _weekday_for(e.get('post_date'))
        if e.get('dr_cr') == 'Cr':
            enriched['credit_amt'] = e['amount']
            enriched['debit_amt'] = None
        else:
            enriched['credit_amt'] = None
            enriched['debit_amt'] = e['amount']
        out.append(enriched)
    return out


GL_DESCRIPTIONS = {
    '3493': 'Outward RTGS/NEFT',
    '3496': 'Inward RTGS/NEFT',
    '345051': 'IMPS/UPI/POS/ACH/NACH',
}


@app.get('/runs/{run_id}/export/{fmt}')
def export_run_view(request: Request, run_id: int, fmt: str, gl: str = '3493', status: str = 'all'):
    user, run, detail = _load_run_detail_or_403(request, run_id)
    if isinstance(detail, RedirectResponse):
        return detail

    gl = gl if gl in ('3493', '3496', '345051') else '3493'
    entries = detail['per_gl'].get(gl, [])
    if status == 'matched':
        entries = [e for e in entries if e['status'].startswith('Matched')]
    elif status == 'pending':
        entries = [e for e in entries if not e['status'].startswith('Matched')
                   and 'Not Applicable' not in e['status'] and not e['status'].startswith('Excluded')]
    elif status == 'excluded':
        entries = [e for e in entries if e['status'].startswith('Excluded')]
    entries = _enrich_entries(entries)

    gl_desc = GL_DESCRIPTIONS[gl]
    label = f"{run['run_label']} - {gl} ({status})"
    base_filename = f"reconciliation_{gl}_{status}_{run['created_at'][:10]}"

    if fmt == 'csv':
        content = export_csv(entries, gl, gl_desc, label)
        return Response(content, media_type='text/csv',
                         headers={'Content-Disposition': f'attachment; filename="{base_filename}.csv"'})
    elif fmt == 'xlsx':
        content = export_xlsx(entries, gl, gl_desc, label)
        return Response(content, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         headers={'Content-Disposition': f'attachment; filename="{base_filename}.xlsx"'})
    elif fmt == 'pdf':
        content = export_pdf(entries, gl, gl_desc, label)
        return Response(content, media_type='application/pdf',
                         headers={'Content-Disposition': f'attachment; filename="{base_filename}.pdf"'})
    else:
        raise HTTPException(404, 'Unknown export format')


@app.get('/runs/{run_id}/manual-clear', response_class=HTMLResponse)
def manual_clear_form(request: Request, run_id: int, gl: str = '3493', idx: int = 0):
    user, run, detail = _load_run_detail_or_403(request, run_id)
    if isinstance(detail, RedirectResponse):
        return detail

    gl = gl if gl in ('3493', '3496', '345051') else '3493'
    entries = detail['per_gl'].get(gl, [])
    if idx < 0 or idx >= len(entries):
        raise HTTPException(404, 'Entry not found')
    entry = entries[idx]

    return templates.TemplateResponse(request, 'manual_clear.html', {
        'user': user, 'run': run, 'active_gl': gl, 'gl_desc': GL_DESCRIPTIONS[gl],
        'entry': entry, 'idx': idx, 'active': 'history',
        'is_excluded': entry['status'].startswith('Excluded'),
    })


@app.post('/runs/{run_id}/manual-clear')
async def manual_clear_submit(request: Request, run_id: int, gl: str = Form(...), idx: int = Form(...),
                                reason: str = Form(...)):
    user = require_login(request)
    if not user:
        raise HTTPException(401, 'Not logged in')
    run = db.get_run(run_id)
    if not run:
        raise HTTPException(404, 'Run not found')
    if not user['is_admin'] and run['user_id'] != user['id']:
        raise HTTPException(403, 'Not your report')
    if not run['detail_path'] or not os.path.exists(run['detail_path']):
        raise HTTPException(404, 'Entry-level detail is not available for this run.')

    gl = gl if gl in ('3493', '3496', '345051') else '3493'
    with open(run['detail_path']) as f:
        detail = json.load(f)

    entries = detail['per_gl'][gl]
    if idx < 0 or idx >= len(entries):
        raise HTTPException(404, 'Entry not found')

    entry = entries[idx]
    timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')
    entry['prior_status'] = entry['status']
    entry['prior_note'] = entry['note']
    entry['status'] = 'Excluded - Manually Reviewed'
    entry['note'] = f"Manually reviewed by {user['display_name']} on {timestamp}: {reason.strip()}"

    with open(run['detail_path'], 'w') as f:
        json.dump(detail, f)

    wb, new_summary = build_workbook_from_detail(detail)
    wb.save(run['report_path'])
    db.update_run_summary(run_id, new_summary)

    return RedirectResponse(f'/runs/{run_id}?gl={gl}&status=pending', status_code=302)


@app.post('/runs/{run_id}/manual-undo')
async def manual_undo_submit(request: Request, run_id: int, gl: str = Form(...), idx: int = Form(...)):
    user = require_login(request)
    if not user:
        raise HTTPException(401, 'Not logged in')
    run = db.get_run(run_id)
    if not run:
        raise HTTPException(404, 'Run not found')
    if not user['is_admin'] and run['user_id'] != user['id']:
        raise HTTPException(403, 'Not your report')
    if not run['detail_path'] or not os.path.exists(run['detail_path']):
        raise HTTPException(404, 'Entry-level detail is not available for this run.')

    gl = gl if gl in ('3493', '3496', '345051') else '3493'
    with open(run['detail_path']) as f:
        detail = json.load(f)

    entries = detail['per_gl'][gl]
    if idx < 0 or idx >= len(entries):
        raise HTTPException(404, 'Entry not found')

    entry = entries[idx]
    if 'prior_status' in entry:
        entry['status'] = entry.pop('prior_status')
        entry['note'] = entry.pop('prior_note', entry['note'])

    with open(run['detail_path'], 'w') as f:
        json.dump(detail, f)

    wb, new_summary = build_workbook_from_detail(detail)
    wb.save(run['report_path'])
    db.update_run_summary(run_id, new_summary)

    return RedirectResponse(f'/runs/{run_id}?gl={gl}&status=excluded', status_code=302)


@app.get('/runs/{run_id}/recheck', response_class=HTMLResponse)
def recheck_form(request: Request, run_id: int, gl: str = '3493'):
    user, run, detail = _load_run_detail_or_403(request, run_id)
    if isinstance(detail, RedirectResponse):
        return detail

    gl = gl if gl in ('3493', '3496', '345051') else '3493'
    entries = detail['per_gl'].get(gl, [])
    pending_entries = [e for e in entries if not e['status'].startswith('Matched') and 'Not Applicable' not in e['status']]

    return templates.TemplateResponse(request, 'recheck.html', {
        'user': user, 'run': run, 'active_gl': gl, 'gl_desc': GL_DESCRIPTIONS[gl],
        'pending_count': len(pending_entries), 'active': 'history',
    })


@app.post('/runs/{run_id}/recheck')
async def recheck_submit(request: Request, run_id: int, gl: str = Form(...), file_hdfc: UploadFile = File(...)):
    user = require_login(request)
    if not user:
        raise HTTPException(401, 'Not logged in')
    run = db.get_run(run_id)
    if not run:
        raise HTTPException(404, 'Run not found')
    if not user['is_admin'] and run['user_id'] != user['id']:
        raise HTTPException(403, 'Not your report')
    if not run['detail_path'] or not os.path.exists(run['detail_path']):
        raise HTTPException(404, 'Entry-level detail is not available for this run.')

    gl = gl if gl in ('3493', '3496', '345051') else '3493'

    run_uuid = uuid.uuid4().hex[:12]
    tmp_path = os.path.join(UPLOADS_DIR, f'{run_uuid}_recheck_{file_hdfc.filename}')
    with open(tmp_path, 'wb') as f:
        shutil.copyfileobj(file_hdfc.file, f)

    try:
        new_hdfc_rows = load_hdfc(tmp_path)
        hdfc_by_amount = {}
        for h in new_hdfc_rows:
            hdfc_by_amount.setdefault(round(h['Amount'], 2), []).append(h)

        with open(run['detail_path']) as f:
            detail = json.load(f)

        entries = detail['per_gl'][gl]
        cleared_count = 0
        for e in entries:
            if e['status'].startswith('Matched') or 'Not Applicable' in e['status']:
                continue  # only re-check genuinely pending entries
            gl_row = rebuild_gl_row_from_entry(e['narration'], e['amount'], e['dr_cr'])
            candidates = find_candidates(gl_row, hdfc_by_amount)
            unused = [(h, reason) for h, reason in candidates if not h['used']]
            if len(unused) == 1:
                h, reason = unused[0]
                h['used'] = True
                e['status'] = 'Matched - Cleared Next Day'
                e['matched_ref'] = h['Reference No']
                e['matched_date'] = h['Transaction Date']
                e['note'] = f"{reason} - cleared next day, matched against uploaded statement ({file_hdfc.filename})"
                cleared_count += 1

        with open(run['detail_path'], 'w') as f:
            json.dump(detail, f)

        wb, new_summary = build_workbook_from_detail(detail)
        wb.save(run['report_path'])
        db.update_run_summary(run_id, new_summary)

        return RedirectResponse(f'/runs/{run_id}?gl={gl}&status=all&cleared={cleared_count}', status_code=302)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


@app.get('/runs/{run_id}', response_class=HTMLResponse)
def view_run(request: Request, run_id: int, gl: str = '3493', status: str = 'all', cleared: int = None):
    user, run, detail = _load_run_detail_or_403(request, run_id)
    if isinstance(detail, RedirectResponse):
        return detail

    gl = gl if gl in ('3493', '3496', '345051') else '3493'
    all_entries = detail['per_gl'].get(gl, [])
    for i, e in enumerate(all_entries):
        e['orig_idx'] = i

    if status == 'matched':
        entries = [e for e in all_entries if e['status'].startswith('Matched')]
    elif status == 'pending':
        entries = [e for e in all_entries if not e['status'].startswith('Matched')
                   and 'Not Applicable' not in e['status'] and not e['status'].startswith('Excluded')]
    elif status == 'excluded':
        entries = [e for e in all_entries if e['status'].startswith('Excluded')]
    else:
        entries = all_entries

    entries = _enrich_entries(entries)

    summary = json.loads(run['summary_json'])
    return templates.TemplateResponse(request, 'run_detail.html', {
        'user': user, 'run': run, 'summary': summary, 'entries': entries,
        'active_gl': gl, 'active_status': status, 'active': 'history', 'cleared': cleared,
    })


@app.get('/api/runs/{run_id}/details')
def api_run_details(request: Request, run_id: int):
    user, run, detail = _load_run_detail_or_403(request, run_id)
    if isinstance(detail, RedirectResponse):
        raise HTTPException(401, 'Not logged in')
    return JSONResponse(detail)


# ---------------------------------------------------------------------------
# Admin: user management
# ---------------------------------------------------------------------------

@app.get('/admin/users', response_class=HTMLResponse)
def admin_users_page(request: Request, error: str = ''):
    user, redirect = require_login_password_ok(request)
    if redirect:
        return redirect
    if not user['is_admin']:
        raise HTTPException(403, 'Admins only')
    users = db.list_users()
    return templates.TemplateResponse(request, 'admin_users.html', {
        'user': user, 'users': users, 'active': 'users', 'error': error,
    })


@app.post('/admin/users')
def admin_create_user(request: Request, username: str = Form(...), password: str = Form(...),
                       display_name: str = Form(''), is_admin: str = Form(None)):
    user = require_login(request)
    if not user:
        return RedirectResponse('/login', status_code=302)
    if not user['is_admin']:
        raise HTTPException(403, 'Admins only')
    if db.get_user_by_username(username.strip().lower()):
        return RedirectResponse('/admin/users?error=exists', status_code=302)
    # admin-created accounts always start with a temporary password that
    # must be changed on first login
    db.create_user(username.strip().lower(), auth.hash_password(password), display_name or username,
                    is_admin=1 if is_admin else 0, must_change_password=1)
    return RedirectResponse('/admin/users', status_code=302)


@app.post('/admin/users/{target_id}/delete')
def admin_delete_user(request: Request, target_id: int):
    user = require_login(request)
    if not user:
        return RedirectResponse('/login', status_code=302)
    if not user['is_admin']:
        raise HTTPException(403, 'Admins only')
    target = db.get_user_by_id(target_id)
    if not target:
        raise HTTPException(404, 'User not found')
    if target_id == user['id']:
        return RedirectResponse('/admin/users?error=You+cannot+delete+your+own+account+while+logged+in', status_code=302)
    if target['is_admin'] and db.count_admins() <= 1:
        return RedirectResponse('/admin/users?error=Cannot+delete+the+last+remaining+admin', status_code=302)
    db.delete_user(target_id)
    return RedirectResponse('/admin/users', status_code=302)


@app.post('/admin/users/{target_id}/reset-password')
def admin_reset_password(request: Request, target_id: int, new_password: str = Form(...)):
    user = require_login(request)
    if not user:
        return RedirectResponse('/login', status_code=302)
    if not user['is_admin']:
        raise HTTPException(403, 'Admins only')
    target = db.get_user_by_id(target_id)
    if not target:
        raise HTTPException(404, 'User not found')
    if len(new_password) < 6:
        return RedirectResponse('/admin/users?error=Password+must+be+at+least+6+characters', status_code=302)
    # a reset always forces that person to choose their own password next
    # time they log in, and clears any lockout at the same time
    db.set_password(target_id, auth.hash_password(new_password), must_change_password=1)
    return RedirectResponse('/admin/users', status_code=302)


@app.post('/admin/users/{target_id}/unlock')
def admin_unlock_user(request: Request, target_id: int):
    user = require_login(request)
    if not user:
        return RedirectResponse('/login', status_code=302)
    if not user['is_admin']:
        raise HTTPException(403, 'Admins only')
    target = db.get_user_by_id(target_id)
    if not target:
        raise HTTPException(404, 'User not found')
    db.unlock_user(target_id)
    return RedirectResponse('/admin/users', status_code=302)


if __name__ == '__main__':
    import uvicorn
    port = int(os.environ.get('PORT', 8000))
    uvicorn.run('app:app', host='0.0.0.0', port=port, reload=False)
