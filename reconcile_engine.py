#!/usr/bin/env python3
"""
RTGS/NEFT/IMPS/UPI/ACH/POS Reconciliation Tool
------------------------------------------------
Reconciles internal GL suspense-head statements (3493 - Outward RTGS/NEFT,
3496 - Inward RTGS/NEFT, 345051 - Other modes: IMPS/UPI/POS/ACH/NACH/AADHAAR/
Returns/CTS) against the HDFC nodal (sub-membership) account statement.

MATCHING PHILOSOPHY: EXACT MATCH ONLY.
  - A GL entry is "Matched" only when exactly one HDFC entry shares the same
    amount AND a shared exact reference code / narration token.
  - Anything ambiguous (0 candidates, 2+ candidates, or amount-only overlap
    without a code match) is left as "Pending" for manual review.
  - Nothing is ever dropped. Every GL row and every HDFC row appears
    somewhere in the output.

Usage:
    python3 reconcile.py <3493.csv> <3496.csv> <345051.csv> <hdfc_main.csv> <output.xlsx>
"""

import csv
import re
import sys
from collections import defaultdict

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def norm_alnum(s):
    """Uppercase, strip everything except letters/digits."""
    return re.sub(r'[^A-Z0-9]', '', (s or '').upper())


def parse_amount(s):
    if s is None:
        return None
    s = s.strip()
    if s == '' or s == '0.00':
        return None
    s = s.replace(',', '').replace('CR', '').replace('DR', '').strip()
    try:
        val = float(s)
        return val if val != 0 else None
    except ValueError:
        return None


DATE_TOKEN_RE = re.compile(r'\d{1,2}[/-]\d{1,2}([/-]\d{2,4})?')
TRF_PREFIX_RE = re.compile(r'^(TO|BY)\s+TRF\s*', re.IGNORECASE)


def normalize_label(text):
    """Build a date-agnostic, prefix-agnostic label key used for matching
    batch-settlement lines (UPI/IMPS/POS/ATM/CTS/ACH/APB) that carry no
    discrete reference number - only a descriptive narration that both the
    GL and the HDFC statement independently generate, sometimes with a
    leading date that can differ by a day due to settlement lag."""
    if not text:
        return ''
    t = TRF_PREFIX_RE.sub('', text)
    t = DATE_TOKEN_RE.sub('', t)
    t = re.sub(r'\bDT:?\b', '', t, flags=re.IGNORECASE)
    return norm_alnum(t)


def extract_codes(raw_text, min_len=10):
    """Extract candidate exact-match numeric codes from the ORIGINAL
    (pre-normalization) text, where hyphens are still visible as internal
    separators within a single code (e.g. '013-3534-00000014'). We match on
    a digit-and-hyphen run bounded by non-digit/non-hyphen characters (space,
    slash, letter), then strip the hyphens out of just that run.

    Doing this on the raw text (instead of a fully-stripped blob) avoids the
    bug where two adjacent reference numbers separated only by a hyphen
    (e.g. 'HDFCR52026082550178267-013353400000014') get fused into one
    unrecognizable digit string once all punctuation is removed first.
    """
    codes = set()
    for m in re.finditer(r'(?<![\d-])[\d-]{6,}(?![\d-])', raw_text or ''):
        digits_only = m.group(0).replace('-', '')
        if len(digits_only) >= min_len:
            codes.add(digits_only)
    return codes


NAME_STRIP_PREFIX_RE = re.compile(r'^(?:(?:NEFT|RTGS)\s+)+', re.IGNORECASE)


GENERIC_NAME_BLOCKLIST = {
    'PHONEPELIMITED', 'PAYTMPAYMENTSSERVICESLTD', 'PAYTMPAYMENTSSERVICESLIMITED',
    'GOOGLEPAY', 'GOOGLEINDIADIGITALSERVICES', 'RAZORPAYSOFTWAREPRIVATELIMITED',
    'CASHFREEPAYMENTSINDIAPVTLTD', 'BILLDESK', 'PAYUPAYMENTSPRIVATELIMITED',
}


def extract_name_key(narration, min_len=6):
    """Pull out the beneficiary/remitter name segment from an RTGS/NEFT
    narration for use as a corroborating (not primary) match signal. RTGS
    narrations are '/'-delimited with the party name as one of the plain-text
    segments (little/no digits), e.g. '099-920102-00000001/PHOENIX IT PARK'
    -> 'PHOENIX IT PARK'. Only used as a tie-breaker on top of an amount
    match (see Rule D) - never on its own - since a name alone can't
    disambiguate two different transactions."""
    best = ''
    for seg in re.split(r'[/\-]+', narration):
        seg = seg.strip()
        # a stray leading "NEFT "/"RTGS " sometimes rides along with the
        # party name inside one '/'-delimited segment (no separator between
        # them) - strip it so it doesn't get baked into the name key.
        seg = NAME_STRIP_PREFIX_RE.sub('', seg).strip()
        if len(seg) < min_len:
            continue
        digit_ratio = sum(c.isdigit() for c in seg) / max(len(seg), 1)
        if digit_ratio > 0.3:
            continue
        upper = seg.upper()
        if upper in ('RTGS', 'NEFT', 'YOUR SELF RTGS', 'YOUR SELF NEFT', 'Y.S. NEFT', 'YS FOR NEFT', 'YS FOR RTGS'):
            continue
        if len(seg) > len(best):
            best = seg
    key = norm_alnum(best) if len(best) >= min_len else None
    # a bare payment-aggregator name (PhonePe/Paytm/etc, with no merchant or
    # transaction-specific text alongside it) is shared by hundreds of
    # unrelated transactions and must never be used as a match key on its
    # own - it only appears here as a short segment because the merchant
    # name happened to be the LONGEST segment in some OTHER row, not this
    # one, so this guard only bites when the aggregator name genuinely is
    # the whole key.
    if key in GENERIC_NAME_BLOCKLIST:
        return None
    return key


def extract_alnum_codes(raw_text, min_len=8):
    """Extract bank-assigned alphanumeric transaction/UTR-style codes such
    as 'AHDSCLGJT041', 'UTGUJ631145K388', 'GHBE91A12480', 'HPCRET41078345R111'.
    These commonly appear as a standalone '/'-delimited token in RTGS/NEFT
    narrations. A token qualifies only if it mixes letters AND digits (so we
    don't pick up plain company-name words) and isn't an IFSC code (IFSC
    identifies a bank branch, not a specific transaction, so matching on it
    alone would create false positives across unrelated transfers to the
    same bank branch)."""
    codes = set()
    IFSC_RE = re.compile(r'^[A-Z]{4}0[A-Z0-9]{6}$')
    for tok in re.findall(r'[A-Za-z0-9]{%d,}' % min_len, raw_text or ''):
        tok_u = tok.upper()
        if IFSC_RE.match(tok_u):
            continue
        has_letter = any(c.isalpha() for c in tok_u)
        has_digit = any(c.isdigit() for c in tok_u)
        if has_letter and has_digit:
            codes.add(tok_u)
    return codes


# ---------------------------------------------------------------------------
# HDFC statement parser
# ---------------------------------------------------------------------------

def load_hdfc(path):
    rows = []
    with open(path, newline='', encoding='utf-8-sig', errors='replace') as f:
        reader = csv.reader(f)
        started = False
        for r in reader:
            if not r:
                continue
            if not started:
                if r[0].strip() == 'Transaction Date':
                    started = True
                continue
            if len(r) < 8:
                continue
            txn_date, desc, amt, dc, ref, val_date, branch, bal = r[:8]
            amount = parse_amount(amt)
            if amount is None:
                continue
            norm_desc = norm_alnum(desc)
            norm_ref = norm_alnum(ref)
            rows.append({
                'row_id': len(rows),
                'Transaction Date': txn_date,
                'Description': desc,
                'Amount': amount,
                'Dr/Cr': dc.strip(),
                'Reference No': ref.strip(),
                'Value Date': val_date,
                'Branch': branch,
                'Running Balance': bal,
                'norm_desc': norm_desc,
                'norm_ref': norm_ref,
                # fully-concatenated haystack (all separators stripped) used
                # for SUBSTRING containment checks of a GL-side embedded code.
                # Deliberately NOT split back into discrete codes here, since
                # two adjacent reference numbers separated by a single hyphen
                # in the raw text (e.g. 'HDFCR...267-013353400000014') would
                # otherwise be indistinguishable from one long code.
                'full_blob': norm_desc + norm_ref,
                'label': normalize_label(desc),
                'name_key': extract_name_key(desc),
                'used': False,
            })
    return rows


# ---------------------------------------------------------------------------
# GL suspense-head parser (3493 / 3496 / 345051 style export)
# ---------------------------------------------------------------------------

INTERNAL_TRANSFER_PATTERNS = [
    r'^AMT TRF (TO|FROM) \d+',
    r'^\d+ CTS CLG CHEQUE RETURN SESSION',  # keep external clearing returns visible but tag separately if needed
]


def is_internal_sweep(narration_upper):
    # These are GL-to-GL end-of-day sweep entries (e.g. "AMT TRF TO 345051",
    # "amt trf from 3493") - they move balances between the bank's own GL
    # heads and will never appear in the HDFC bank statement, so they are
    # not part of bank reconciliation at all.
    return bool(re.search(r'AMT\s+TRF\s+(TO|FROM)\s+\d+', narration_upper))


def load_gl(path):
    rows = []
    with open(path, newline='', encoding='utf-8-sig', errors='replace') as f:
        reader = csv.reader(f)
        started = False
        for r in reader:
            if not r:
                continue
            if not started:
                if len(r) >= 3 and r[0].strip() == 'Textbox21' and r[1].strip() == 'POSTDATE':
                    started = True
                continue
            if len(r) < 5:
                continue
            post_date, value_date, narration, debit, credit = r[:5]
            debit_amt = parse_amount(debit)
            credit_amt = parse_amount(credit)
            if debit_amt is None and credit_amt is None:
                continue
            if debit_amt is not None:
                amount, side = debit_amt, 'Dr'
            else:
                amount, side = credit_amt, 'Cr'

            narration_clean = narration.strip()
            internal = is_internal_sweep(narration_clean.upper())

            norm_narr = norm_alnum(narration_clean)
            # strip a leading quote character some exports leave behind
            narration_for_label = narration_clean.lstrip('"\' ')
            batch_label = normalize_label(narration_for_label)

            rows.append({
                'row_id': len(rows),
                'Post Date': post_date,
                'Value Date': value_date,
                'Narration': narration_clean,
                'Amount': amount,
                'Dr/Cr': side,
                'Internal Sweep': internal,
                'norm_narr': norm_narr,
                'codes': extract_codes(narration_clean) | extract_alnum_codes(narration_clean),
                'name_key': extract_name_key(narration_clean),
                'batch_label': batch_label,
                'status': None,
                'matched_hdfc_row': None,
                'note': None,
            })
    return rows


# ---------------------------------------------------------------------------
# Matching engine
# ---------------------------------------------------------------------------

def find_candidates(gl_row, hdfc_by_amount):
    amount = gl_row['Amount']
    pool = [h for h in hdfc_by_amount.get(round(amount, 2), []) if not h['used']]
    if not pool:
        return []

    candidates = []
    gl_codes = gl_row['codes']

    for h in pool:
        matched_reason = None
        # Rule A: HDFC's own Reference No (a real UTR/reference, len>=8)
        # appears verbatim inside the GL narration.
        if h['norm_ref'] and len(h['norm_ref']) >= 8 and h['norm_ref'] in gl_row['norm_narr']:
            matched_reason = 'Reference No. match'

        # Rule B: a long digit/alnum code embedded in the GL narration
        # appears verbatim as a substring inside the HDFC narration/reference
        # blob. Substring (not set-equality) is essential: HDFC often shows
        # two reference numbers joined by a single hyphen, so the GL's
        # isolated code still appears as a contiguous run inside the longer
        # blob even though it can't be cleanly re-split back out of it.
        if matched_reason is None and any(code in h['full_blob'] for code in gl_codes):
            matched_reason = 'Embedded transaction code match'

        # Rule C: batch-settlement label match (IMPS/UPI/POS/ATM/CTS/ACH/APB
        # lines that carry no reference number at all) - a date-agnostic,
        # prefix-agnostic exact text match, since both the GL narration and
        # the HDFC description independently render the same settlement
        # batch label, sometimes with a leading date that can differ by a
        # day due to settlement lag. Each check below only fires if the
        # PREVIOUS rule found nothing - a truthy label on both sides that
        # simply doesn't match must still fall through to Rule D, not stop
        # the search.
        if matched_reason is None and gl_row['batch_label'] and len(gl_row['batch_label']) >= 8 and h['label']:
            if gl_row['batch_label'] == h['label']:
                matched_reason = 'Batch label match (exact)'
            elif gl_row['batch_label'] in h['label'] or h['label'] in gl_row['batch_label']:
                matched_reason = 'Batch label match (contains)'

        # Rule D: no reference code matched on either side, but exactly one
        # HDFC entry at this amount contains the GL row's beneficiary /
        # remitter name verbatim (checked in both directions against the
        # FULL opposite narration - never name-key-vs-name-key, since two
        # short specific keys can spuriously overlap in ways a short generic
        # key like a payment app's own name should not be trusted to
        # disambiguate). Still an exact (substring) match, just a weaker
        # corroborating signal than a reference code, so it's tried last.
        if matched_reason is None and gl_row.get('name_key') and gl_row['name_key'] in h['norm_desc']:
            matched_reason = 'Beneficiary/remitter name match'
        elif matched_reason is None and h.get('name_key') and h['name_key'] in gl_row['norm_narr']:
            matched_reason = 'Beneficiary/remitter name match'

        if matched_reason:
            candidates.append((h, matched_reason))
    return candidates


def reconcile(gl_rows, hdfc_rows):
    hdfc_by_amount = defaultdict(list)
    for h in hdfc_rows:
        hdfc_by_amount[round(h['Amount'], 2)].append(h)

    pending_multi = []  # (gl_row, candidate_list) needing phase-2 resolution

    for gl_row in gl_rows:
        if gl_row['Internal Sweep']:
            gl_row['status'] = 'Not Applicable - Internal GL Sweep'
            gl_row['note'] = 'End-of-day transfer between internal GL heads; never hits the HDFC statement.'
            continue

        candidates = find_candidates(gl_row, hdfc_by_amount)

        if len(candidates) == 1:
            h, reason = candidates[0]
            h['used'] = True
            gl_row['status'] = 'Matched'
            gl_row['matched_hdfc_row'] = h['row_id']
            gl_row['note'] = reason
        elif len(candidates) == 0:
            gl_row['status'] = 'Pending - Missing in Bank Statement'
            gl_row['note'] = 'No HDFC entry found with matching amount and reference/code.'
        else:
            gl_row['status'] = 'Pending - Multiple Possible Matches'
            gl_row['note'] = f'{len(candidates)} HDFC entries share this amount and code; needs manual review.'
            pending_multi.append((gl_row, candidates))

    # ---- Phase 2: resolve true duplicate groups by chronological order ----
    # If several GL rows carry the IDENTICAL set of HDFC candidates (same
    # amount, same embedded reference code - i.e. genuinely repeated
    # transactions to the same beneficiary), and the group has exactly as
    # many GL rows as candidate HDFC rows, this is not really ambiguous -
    # it's a set of true duplicates. Pair them up in the order each side
    # occurs in its own file (1st GL <-> 1st HDFC, 2nd <-> 2nd, ...). This
    # is still an exact match (every pairing shares amount + reference
    # code) with a deterministic, non-guessed tie-break - never crossing
    # into a different amount or a different code.
    groups = defaultdict(list)
    for gl_row, candidates in pending_multi:
        key = tuple(sorted(h['row_id'] for h, _ in candidates))
        groups[key].append((gl_row, candidates))

    for key, members in groups.items():
        # only the GL rows still unresolved and only unused HDFC candidates
        cand_pool = [h for h in (hdfc_rows[i] for i in key) if not h['used']]
        if len(members) == len(cand_pool) and len(members) > 1:
            members_sorted = sorted(members, key=lambda m: m[0]['row_id'])
            cand_pool_sorted = sorted(cand_pool, key=lambda h: h['row_id'])
            for (gl_row, candidates), h in zip(members_sorted, cand_pool_sorted):
                reason = candidates[0][1]
                h['used'] = True
                gl_row['status'] = 'Matched'
                gl_row['matched_hdfc_row'] = h['row_id']
                gl_row['note'] = (f'{reason} - duplicate group of {len(members)} identical '
                                   f'amount+reference entries, paired in chronological order')

    return hdfc_rows


def rebuild_gl_row_from_entry(narration, amount, dr_cr):
    """Reconstruct the minimal fields find_candidates() needs, starting only
    from a previously-saved entry's narration/amount/side (used by the
    're-check pending against a later statement' flow, where we no longer
    have the original rich in-memory GL row - just what was serialized to
    the run's detail JSON)."""
    norm_narr = norm_alnum(narration)
    narration_for_label = narration.lstrip('"\' ')
    return {
        'Narration': narration,
        'Amount': amount,
        'Dr/Cr': dr_cr,
        'Internal Sweep': False,
        'norm_narr': norm_narr,
        'codes': extract_codes(narration) | extract_alnum_codes(narration),
        'batch_label': normalize_label(narration_for_label),
        'name_key': extract_name_key(narration),
    }


if __name__ == '__main__':
    pass
