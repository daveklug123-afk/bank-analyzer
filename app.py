import os, json, re, io, anthropic, pickle, threading, uuid, time
from flask import Flask, request, jsonify, send_file, render_template, session, redirect, url_for
from werkzeug.utils import secure_filename
import pdfplumber, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'mca-analyzer-secret-2026')
app.config['UPLOAD_FOLDER'] = os.environ.get('UPLOAD_FOLDER', '/tmp/uploads')
app.config['HISTORY_FOLDER'] = os.environ.get('HISTORY_FOLDER', '/tmp/history')
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024
ALLOWED_EXTENSIONS = {'pdf', 'csv', 'txt'}
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['HISTORY_FOLDER'], exist_ok=True)

USERS = {
    os.environ.get('USERNAME1', 'dave'): os.environ.get('PASSWORD1', 'mca2026'),
    os.environ.get('USERNAME2', 'admin'): os.environ.get('PASSWORD2', 'analyze2026'),
}

# ── In-memory job store ──────────────────────────────────────────────────────
# { job_id: { status, progress, message, data, excel_bytes, error } }
JOBS = {}
JOBS_LOCK = threading.Lock()

def job_set(job_id, **kwargs):
    with JOBS_LOCK:
        if job_id not in JOBS:
            JOBS[job_id] = {}
        JOBS[job_id].update(kwargs)

def job_get(job_id):
    with JOBS_LOCK:
        return dict(JOBS.get(job_id, {}))

def run_analysis_job(job_id, combined_text, company_name, entry_id, existing_entry):
    """Runs in a background thread — no Gunicorn timeout applies."""
    try:
        # Step 1: Parse
        job_set(job_id, status="running", progress=20,
                message="Pass 1 of 2 — Extracting data from statement...")
        new_data = parse_with_claude(combined_text, company_name)
        new_data = sanitize_data(new_data)

        # Step 2: Verify
        job_set(job_id, status="running", progress=55,
                message="Pass 2 of 2 — Running independent verification check...")
        try:
            verify_flags = verify_with_claude(combined_text, new_data)
            existing_flags = new_data.get("review_flags", [])
            new_data["review_flags"] = verify_flags + existing_flags
        except Exception as e:
            new_data.setdefault("review_flags", []).insert(0, {
                "type": "VERIFY_ERROR",
                "field": "verification_pass",
                "value": "Verification did not run",
                "confidence": 0.0,
                "reason": "Second-pass verification error: {}".format(str(e)),
                "month": "ALL"
            })

        # Step 3: Merge with existing if needed
        if existing_entry:
            new_data = merge_data(existing_entry['data'], new_data)

        # Step 4: Build Excel
        job_set(job_id, status="running", progress=85,
                message="Building Excel output...")
        excel = build_excel(new_data)
        excel_bytes = excel.read()

        # Step 5: Save history
        cn = new_data.get("company_name", "Unknown")
        save_history(cn, new_data, excel_bytes)
        safe = re.sub(r'[^\w\s-]', '', cn).strip().replace(' ', '_')

        job_set(job_id, status="done", progress=100,
                message="Complete",
                data=new_data,
                excel_bytes=excel_bytes,
                filename=safe + "_analysis.xlsx")

    except Exception as e:
        job_set(job_id, status="error", progress=0,
                message="Analysis failed",
                error=str(e))

def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def allowed_file(f): return '.' in f and f.rsplit('.',1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_pdf(path):
    text_parts = []
    try:
        with pdfplumber.open(path) as pdf:
            total = len(pdf.pages)
            if total <= 12:
                pages_to_read = list(range(total))
            else:
                pages_to_read = list(range(8)) + list(range(max(8, total-4), total))
            for i in pages_to_read:
                try:
                    t = pdf.pages[i].extract_text()
                    if t:
                        text_parts.append(t[:3000])
                except:
                    pass
    except Exception as e:
        return ""
    return "\n".join(text_parts)

def extract_text(path):
    ext = path.rsplit('.',1)[1].lower()
    if ext == 'pdf': return extract_text_from_pdf(path)
    with open(path,'r',encoding='utf-8',errors='ignore') as f: return f.read()

def calc_monthly(amount, frequency):
    freq = (frequency or 'weekly').lower()
    if 'daily' in freq: return amount * 22
    if 'bi' in freq: return amount * 2
    if 'monthly' in freq: return amount * 1
    return amount * 4

def save_history(company_name, data, excel_bytes):
    safe = re.sub(r'[^\w\s-]','',company_name).strip().replace(' ','_')
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    entry_id = "{}_{}".format(safe, ts)
    path = os.path.join(app.config['HISTORY_FOLDER'], entry_id + '.pkl')
    with open(path, 'wb') as f:
        pickle.dump({'id':entry_id,'company_name':company_name,'data':data,'excel':excel_bytes,'timestamp':ts}, f)
    return entry_id

def load_history():
    entries = []
    folder = app.config['HISTORY_FOLDER']
    for fname in sorted(os.listdir(folder), reverse=True)[:10]:
        if fname.endswith('.pkl'):
            try:
                with open(os.path.join(folder, fname), 'rb') as f:
                    e = pickle.load(f)
                    entries.append({'id':e['id'],'company_name':e['company_name'],'timestamp':e['timestamp']})
            except: pass
    return entries

def sanitize_data(obj):
    if isinstance(obj, dict):
        return {k: sanitize_data(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [sanitize_data(i) for i in obj]
    elif isinstance(obj, str):
        result = []
        for ch in obj:
            cp = ord(ch)
            if cp < 0x20 and cp not in (0x09, 0x0a, 0x0d): continue
            if 0xD800 <= cp <= 0xDFFF: continue
            if 0xFFFE <= cp <= 0xFFFF: continue
            result.append(ch)
        s = ''.join(result)
        s = s.replace('–','-').replace('—','-')
        s = s.replace('\u2018',"'").replace('\u2019',"'")
        s = s.replace('"','"').replace('"','"')
        s = s.replace('•','*').replace('\u00a0',' ')
        s = s.replace('…','...').replace('−','-')
        s = s.replace('·','*').replace('●','*')
        return s
    return obj

def load_entry(entry_id):
    path = os.path.join(app.config['HISTORY_FOLDER'], entry_id + '.pkl')
    if not os.path.exists(path): return None
    with open(path, 'rb') as f:
        entry = pickle.load(f)
    if 'data' in entry:
        entry['data'] = sanitize_data(entry['data'])
    return entry

def parse_with_claude(raw_text, company_name=""):
    client = anthropic.Anthropic()
    cn = company_name if company_name else 'auto-detect'
    prompt = (
        "You are an expert MCA (Merchant Cash Advance) underwriter analyzing bank statements.\n"
        "You are the LENDER deciding whether to advance money to this business.\n"
        "Return ONLY valid JSON, no markdown, no explanation.\n\n"
        "BANK STATEMENT TEXT:\n"
        + raw_text[:80000] +
        "\n\n"
        "=== CRITICAL EXTRACTION RULES ===\n\n"
        "RULE 1 - CURRENT POSITIONS (MCA lenders taking recurring ACH debits):\n"
        "- Look in the transaction history for recurring ACH debits labeled 'Business to Business ACH Debit'\n"
        "- amount = the EXACT per-payment dollar amount (e.g. $450.00, not $9,900)\n"
        "- frequency = how often they debit:\n"
        "  * If the same company debits EVERY SINGLE BUSINESS DAY = 'daily'\n"
        "  * If they debit once per week = 'weekly'\n"
        "  * If they debit every other week = 'bi-weekly'\n"
        "  * If they debit once per month = 'monthly'\n"
        "  COUNT the actual debits in the statement to determine frequency\n"
        "- DO NOT multiply the amount by frequency - report the raw per-payment amount\n\n"
        "RULE 2 - TRUE DEPOSITS (only real business revenue):\n"
        "INCLUDE:\n"
        "  - POS/credit card processor deposits (Stripe, Lightspeed, Square, Clover, Synchrony Mtot Dep)\n"
        "  - eDeposit IN Branch (cash deposits)\n"
        "  - Mobile Deposits\n"
        "  - ACH credits from real customers/vendors\n"
        "  - Interest payments\n"
        "EXCLUDE (these are NOT real revenue):\n"
        "  - MCA funding credits\n"
        "  - Incoming wire or ACH credit from known MCA lenders\n"
        "  - Online transfers FROM personal/other business accounts\n"
        "  - Book transfers between own accounts\n"
        "  - Returned item credits\n\n"
        "RULE 3 - LEVERAGE PERCENTAGE:\n"
        "  leverage_pct = (sum of all monthly MCA payment amounts / true monthly deposits) * 100\n\n"
        "RULE 4 - NSF / RETURNED ITEMS:\n"
        "  nsf_count = number of items in 'Items returned unpaid' section\n"
        "  od_count = number of 'Overdraft Fee' charges\n\n"
        "RULE 5 - MONTHS:\n"
        "  Extract EVERY statement period found.\n"
        "  Most recent partial month = is_mtd: true\n\n"
        "RULE 6 - TRANSACTION ACCOUNTING (CRITICAL):\n"
        "  Every single transaction in the statement MUST be accounted for.\n"
        "  For each transaction you process:\n"
        "    - It must land in one of: true_deposits, excluded_deposits, current_positions, funding_events, or other_debits\n"
        "    - If a transaction does not clearly fit any known category, add it to 'unclassified_transactions'\n"
        "    - Do NOT silently ignore any transaction\n\n"
        "RULE 7 - CONFIDENCE SCORES (CRITICAL):\n"
        "  For every extracted field, assign a confidence score 0.0 to 1.0:\n"
        "    1.0 = explicitly stated in the document, no ambiguity\n"
        "    0.9 = very clear, minor formatting inference\n"
        "    0.8 = reasonably clear but required some interpretation\n"
        "    0.7 or below = uncertain, ambiguous, or inferred from limited data\n"
        "  Any field with confidence < 1.0 MUST include a 'reason' string explaining WHY it's not 100%.\n"
        "  Be strict - only give 1.0 if the value is explicitly written in the document.\n\n"
        "Return this exact JSON structure:\n"
        '{\n'
        '  "company_name": "string",\n'
        '  "company_name_confidence": 1.0,\n'
        '  "account_number_last4": "string",\n'
        '  "account_number_confidence": 1.0,\n'
        '  "num_bank_accounts": 1,\n'
        '  "offer_decline": "DECLINE",\n'
        '  "holdback_pct": 0.0,\n'
        '  "leverage_pct": 0.0,\n'
        '  "leverage_pct_confidence": 0.9,\n'
        '  "leverage_pct_confidence_reason": "calculated from extracted values",\n'
        '  "sos_info": "",\n'
        '  "court_search_notes": "",\n'
        '  "account_notes": [],\n'
        '  "current_positions": [\n'
        '    {\n'
        '      "lender": "name",\n'
        '      "amount": 0.0,\n'
        '      "amount_confidence": 1.0,\n'
        '      "frequency": "daily/weekly/bi-weekly/monthly",\n'
        '      "frequency_confidence": 0.9,\n'
        '      "frequency_confidence_reason": "counted 18 debits across 22 business days",\n'
        '      "notes": ""\n'
        '    }\n'
        '  ],\n'
        '  "months": [\n'
        '    {\n'
        '      "month_label": "Mon-YY",\n'
        '      "period": "MM/DD to MM/DD",\n'
        '      "is_mtd": false,\n'
        '      "total_deposits": 0.0,\n'
        '      "total_deposits_confidence": 1.0,\n'
        '      "true_deposits": 0.0,\n'
        '      "true_deposits_confidence": 0.9,\n'
        '      "true_deposits_confidence_reason": "",\n'
        '      "true_deposit_exclusions": "",\n'
        '      "neg_days": 0,\n'
        '      "nsf_count": 0,\n'
        '      "nsf_confidence": 1.0,\n'
        '      "od_count": 0,\n'
        '      "num_transactions": 0,\n'
        '      "adb": 0.0,\n'
        '      "adb_confidence": 1.0,\n'
        '      "days_below_1000": 0,\n'
        '      "days_below_1000_confidence": 0.8,\n'
        '      "days_below_1000_confidence_reason": "inferred from daily balance table",\n'
        '      "funding_events": [{"funder": "name", "amount": 0.0, "date": "MM/DD"}],\n'
        '      "unclassified_transactions": [\n'
        '        {\n'
        '          "date": "MM/DD",\n'
        '          "description": "raw transaction description",\n'
        '          "amount": 0.0,\n'
        '          "direction": "credit/debit",\n'
        '          "reason_unclassified": "does not match any known category rule"\n'
        '        }\n'
        '      ],\n'
        '      "notes": ""\n'
        '    }\n'
        '  ],\n'
        '  "balance_reconciliation": {\n'
        '    "opening_balance": 0.0,\n'
        '    "closing_balance": 0.0,\n'
        '    "stated_total_deposits": 0.0,\n'
        '    "stated_total_withdrawals": 0.0,\n'
        '    "calculated_closing_balance": 0.0,\n'
        '    "reconciles": true,\n'
        '    "discrepancy": 0.0\n'
        '  }\n'
        '}\n\n'
        "Company name if provided: " + cn
    )
    msg = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=8000,
        messages=[{"role":"user","content":prompt}]
    )
    raw = msg.content[0].text.strip()
    raw = re.sub(r'^```json\s*','',raw)
    raw = re.sub(r'^```\s*','',raw)
    raw = re.sub(r'\s*```$','',raw)
    data = json.loads(raw)

    def clean_json(obj):
        if isinstance(obj, dict):
            return {k: clean_json(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [clean_json(i) for i in obj]
        elif isinstance(obj, str):
            result = []
            for ch in obj:
                cp = ord(ch)
                if cp < 0x20 and cp not in (0x09, 0x0a, 0x0d): continue
                if 0xD800 <= cp <= 0xDFFF: continue
                if 0xFFFE <= cp <= 0xFFFF: continue
                result.append(ch)
            s = ''.join(result)
            s = s.replace('\u2013','-').replace('\u2014','-')
            s = s.replace('\u2018',"'").replace('\u2019',"'")
            s = s.replace('\u201c','"').replace('\u201d','"')
            s = s.replace('\u2022','*').replace('\u00a0',' ')
            return s
        return obj
    data = clean_json(data)

    # Calculate monthly totals
    total = 0
    for pos in data.get("current_positions", []):
        monthly = calc_monthly(pos.get("amount", 0), pos.get("frequency", "weekly"))
        pos["monthly_amount"] = monthly
        total += monthly
    data["total_current_positions"] = total

    if not data.get("leverage_pct") and data.get("months"):
        full_months = [m for m in data["months"] if not m.get("is_mtd")]
        if full_months:
            latest = full_months[0]
            true_dep = latest.get("true_deposits", 0)
            if true_dep > 0:
                data["leverage_pct"] = round((total / true_dep) * 100, 2)

    # Build the review flags list - collect everything below 100% confidence
    # or unclassified transactions
    review_flags = []

    # Check top-level field confidences
    top_level_confidence_fields = [
        ("company_name", "company_name_confidence", None),
        ("account_number_last4", "account_number_confidence", None),
        ("leverage_pct", "leverage_pct_confidence", "leverage_pct_confidence_reason"),
    ]
    for field, conf_key, reason_key in top_level_confidence_fields:
        conf = data.get(conf_key, 1.0)
        if conf is not None and float(conf) < 1.0:
            reason = data.get(reason_key, "") if reason_key else ""
            review_flags.append({
                "type": "LOW_CONFIDENCE",
                "field": field,
                "value": str(data.get(field, "")),
                "confidence": conf,
                "reason": reason,
                "month": "—"
            })

    # Check position-level confidences
    for pos in data.get("current_positions", []):
        lender = pos.get("lender", "unknown")
        for field, conf_key, reason_key in [
            ("amount", "amount_confidence", None),
            ("frequency", "frequency_confidence", "frequency_confidence_reason"),
        ]:
            conf = pos.get(conf_key, 1.0)
            if conf is not None and float(conf) < 1.0:
                reason = pos.get(reason_key, "") if reason_key else ""
                review_flags.append({
                    "type": "LOW_CONFIDENCE",
                    "field": "position[{}].{}".format(lender, field),
                    "value": str(pos.get(field, "")),
                    "confidence": conf,
                    "reason": reason,
                    "month": "—"
                })

    # Check month-level confidences and unclassified transactions
    for m in data.get("months", []):
        label = m.get("month_label", "?")
        month_confidence_fields = [
            ("total_deposits", "total_deposits_confidence", None),
            ("true_deposits", "true_deposits_confidence", "true_deposits_confidence_reason"),
            ("nsf_count", "nsf_confidence", None),
            ("adb", "adb_confidence", None),
            ("days_below_1000", "days_below_1000_confidence", "days_below_1000_confidence_reason"),
        ]
        for field, conf_key, reason_key in month_confidence_fields:
            conf = m.get(conf_key, 1.0)
            if conf is not None and float(conf) < 1.0:
                reason = m.get(reason_key, "") if reason_key else ""
                review_flags.append({
                    "type": "LOW_CONFIDENCE",
                    "field": field,
                    "value": str(m.get(field, "")),
                    "confidence": conf,
                    "reason": reason,
                    "month": label
                })

        # Collect unclassified transactions
        for txn in m.get("unclassified_transactions", []):
            review_flags.append({
                "type": "UNCLASSIFIED_TXN",
                "field": "transaction",
                "value": "{} | {} | ${:,.2f}".format(
                    txn.get("date","?"),
                    txn.get("description","?"),
                    float(txn.get("amount", 0))
                ),
                "confidence": 0.0,
                "reason": txn.get("reason_unclassified", "Not matched to any category rule"),
                "month": label
            })

    # Balance reconciliation check
    recon = data.get("balance_reconciliation", {})
    if recon and not recon.get("reconciles", True):
        disc = recon.get("discrepancy", 0)
        review_flags.append({
            "type": "BALANCE_MISMATCH",
            "field": "balance_reconciliation",
            "value": "Discrepancy: ${:,.2f}".format(abs(float(disc))),
            "confidence": 0.0,
            "reason": "Opening balance + deposits - withdrawals does not equal stated closing balance",
            "month": "ALL"
        })

    data["review_flags"] = review_flags
    data["balance_reconciliation"] = recon

    return data

def verify_with_claude(raw_text, extracted_data):
    """
    Second independent Claude pass. Reads the raw statement text and the
    already-extracted JSON side by side, then looks for contradictions,
    missed positions, wrong totals, miscounted debits, or anything suspicious.
    Returns a list of verification flags that get merged into review_flags.
    """
    client = anthropic.Anthropic()

    # Compact summary of what was extracted so far
    summary_lines = []
    summary_lines.append("COMPANY: {}".format(extracted_data.get("company_name", "?")))
    summary_lines.append("LEVERAGE: {}%".format(extracted_data.get("leverage_pct", 0)))

    positions = extracted_data.get("current_positions", [])
    summary_lines.append("POSITIONS FOUND ({}):".format(len(positions)))
    for p in positions:
        summary_lines.append("  - {} | ${} | {} | monthly=${:.2f}".format(
            p.get("lender","?"), p.get("amount",0),
            p.get("frequency","?"), p.get("monthly_amount",0)
        ))

    months = extracted_data.get("months", [])
    summary_lines.append("MONTHS FOUND ({}):".format(len(months)))
    for m in months:
        summary_lines.append("  - {} | total_deposits=${} | true_deposits=${} | adb=${} | nsf={} | od={} | days_below_1000={}".format(
            m.get("month_label","?"),
            m.get("total_deposits",0),
            m.get("true_deposits",0),
            m.get("adb",0),
            m.get("nsf_count",0),
            m.get("od_count",0),
            m.get("days_below_1000",0)
        ))
        for fe in m.get("funding_events", []):
            summary_lines.append("    funding: {} ${} on {}".format(
                fe.get("funder","?"), fe.get("amount",0), fe.get("date","?")
            ))

    recon = extracted_data.get("balance_reconciliation", {})
    if recon:
        summary_lines.append("RECONCILIATION: reconciles={} discrepancy={}".format(
            recon.get("reconciles", "?"), recon.get("discrepancy", 0)
        ))

    extracted_summary = "\n".join(summary_lines)

    prompt = (
        "You are a SECOND independent MCA underwriter doing a QA review.\n"
        "The FIRST analyst already extracted data from the bank statement. Your job is to VERIFY their work.\n"
        "Be skeptical. Find errors, omissions, and contradictions.\n"
        "Return ONLY valid JSON — no markdown, no explanation.\n\n"
        "=== WHAT THE FIRST ANALYST EXTRACTED ===\n"
        + extracted_summary +
        "\n\n=== ORIGINAL BANK STATEMENT TEXT ===\n"
        + raw_text[:60000] +
        "\n\n=== YOUR VERIFICATION TASKS ===\n\n"
        "TASK 1 — POSITION COUNT CHECK:\n"
        "  Scan the statement for ALL recurring ACH debits. Count unique payees with 2+ debits.\n"
        "  If you find a recurring debtor NOT in the extracted positions list, flag it.\n"
        "  If a position's frequency looks wrong (e.g., counted as daily but only debited 8 times), flag it.\n"
        "  If a position's per-payment amount looks wrong, flag it.\n\n"
        "TASK 2 — DEPOSIT TOTAL CROSS-CHECK:\n"
        "  For each month, verify the extracted total_deposits matches the statement's summary box.\n"
        "  Verify true_deposits is reasonable (total minus known MCA fundings).\n"
        "  Flag any month where the numbers seem off by more than $100.\n\n"
        "TASK 3 — BALANCE CHECK:\n"
        "  Verify: opening_balance + total_deposits - total_withdrawals - checks - fees = closing_balance.\n"
        "  Flag if it doesn't reconcile within $1.00.\n\n"
        "TASK 4 — NSF / NEGATIVE DAY CHECK:\n"
        "  Scan the daily balance table. Count days where balance dropped below $1,000.\n"
        "  Check for any 'NSF', 'returned item', 'overdraft fee' entries.\n"
        "  Flag if the extracted counts don't match what you see.\n\n"
        "TASK 5 — SUSPICIOUS PATTERNS:\n"
        "  Flag any of these if found:\n"
        "  - Round-trip transfers (money in and out to same entity same day or within 2 days)\n"
        "  - Unusually large single deposits that could be MCA funding not captured\n"
        "  - Payroll wire amounts that seem inconsistent month to month (possible double-counting)\n"
        "  - Any wire OUT over $10,000 not explained by identified positions\n"
        "  - Any wire IN over $5,000 not explained by identified revenue sources or funding events\n\n"
        "TASK 6 — ADB CHECK:\n"
        "  If a daily balance table is present, calculate the average yourself and compare to extracted ADB.\n"
        "  Flag if difference is more than $500.\n\n"
        "Return this exact JSON — include ONLY actual issues found, empty array if everything checks out:\n"
        "{\n"
        "  \"verification_passed\": true,\n"
        "  \"issues\": [\n"
        "    {\n"
        "      \"task\": \"TASK 1\",\n"
        "      \"severity\": \"HIGH/MEDIUM/LOW\",\n"
        "      \"field\": \"exact field name or position lender name\",\n"
        "      \"month\": \"Mon-YY or ALL\",\n"
        "      \"extracted_value\": \"what the first analyst said\",\n"
        "      \"actual_value\": \"what you found in the statement\",\n"
        "      \"description\": \"clear explanation of the discrepancy\"\n"
        "    }\n"
        "  ]\n"
        "}\n"
        "Set verification_passed to false if any HIGH or MEDIUM severity issues exist."
    )

    try:
        msg = client.messages.create(
            model="claude-opus-4-5",
            max_tokens=4000,
            messages=[{"role": "user", "content": prompt}]
        )
        raw = msg.content[0].text.strip()
        raw = re.sub(r'^```json\s*', '', raw)
        raw = re.sub(r'^```\s*', '', raw)
        raw = re.sub(r'\s*```$', '', raw)
        result = json.loads(raw)
    except Exception as e:
        # Verification failed to run — add a single flag noting this
        return [{
            "type": "VERIFY_ERROR",
            "field": "verification_pass",
            "value": "Verification could not complete",
            "confidence": 0.0,
            "reason": "Second-pass Claude verification threw an error: {}".format(str(e)),
            "month": "ALL"
        }]

    issues = result.get("issues", [])
    flags = []
    for issue in issues:
        severity = issue.get("severity", "LOW")
        conf = 0.0 if severity == "HIGH" else (0.5 if severity == "MEDIUM" else 0.75)
        desc = issue.get("description", "")
        extracted = issue.get("extracted_value", "")
        actual = issue.get("actual_value", "")
        detail = desc
        if extracted and actual and extracted != actual:
            detail = "{} | Extracted: {} | Actual: {}".format(desc, extracted, actual)
        flags.append({
            "type": "VERIFY_{}".format(severity),
            "field": issue.get("field", "unknown"),
            "value": extracted,
            "confidence": conf,
            "reason": "[{}] {}".format(issue.get("task", "?"), detail),
            "month": issue.get("month", "?")
        })

    # Add a summary flag if verification failed overall
    if not result.get("verification_passed", True):
        flags.insert(0, {
            "type": "VERIFY_FAILED",
            "field": "SECOND PASS RESULT",
            "value": "{} issue(s) found".format(len(issues)),
            "confidence": 0.0,
            "reason": "Second independent Claude review found contradictions with the source statement. Review all VERIFY flags below.",
            "month": "ALL"
        })
    else:
        flags.insert(0, {
            "type": "VERIFY_PASSED",
            "field": "SECOND PASS RESULT",
            "value": "PASSED",
            "confidence": 1.0,
            "reason": "Second independent Claude review found no contradictions with the source statement.",
            "month": "ALL"
        })

    return flags


def merge_data(existing, new_data):
    existing_labels = {m['month_label'] for m in existing.get('months', [])}
    for m in new_data.get('months', []):
        if m['month_label'] not in existing_labels:
            existing['months'].append(m)
    def month_sort_key(m):
        label = m.get('month_label', '')
        month_map = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
                     'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
        try:
            parts = label.split('-')
            mon = month_map.get(parts[0], 0)
            yr = int(parts[1]) if len(parts) > 1 else 0
            return yr * 100 + mon
        except:
            return 0
    existing['months'].sort(key=month_sort_key, reverse=True)
    existing_lenders = {p['lender'] for p in existing.get('current_positions', [])}
    for p in new_data.get('current_positions', []):
        if p['lender'] not in existing_lenders:
            existing['current_positions'].append(p)
    total = sum(calc_monthly(p.get('amount',0), p.get('frequency','weekly'))
                for p in existing.get('current_positions', []))
    existing['total_current_positions'] = total

    # Merge review flags
    existing_flags = existing.get('review_flags', [])
    new_flags = new_data.get('review_flags', [])
    seen = {(f['field'], f['month'], f['value']) for f in existing_flags}
    for flag in new_flags:
        key = (flag['field'], flag['month'], flag['value'])
        if key not in seen:
            existing_flags.append(flag)
            seen.add(key)
    existing['review_flags'] = existing_flags

    return existing

def build_excel(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Analysis"

    GOLD_PALE="FFF8E1"; LIGHT_YELLOW="FFFF99"; DARK_GOLD="8B6914"
    GREEN_BG="C6EFCE"; GREEN_FG="006100"; RED_BG="FFC7CE"; RED_FG="9C0006"
    BLUE="0070C0"; PURPLE_BG="EAD5F5"; GRAY="F2F2F2"
    NEG_BG="FFC7CE"; OK_BG="C6EFCE"; MONTH_BG="FFD966"
    REVIEW_BG="FFF2CC"; REVIEW_HEADER="FF6600"; UNCLASS_BG="FCE4D6"
    RECON_BG="DEEBF7"

    thin = Side(style='thin')
    def border_all():
        return Border(left=thin,right=thin,top=thin,bottom=thin)

    def w(row,col,value="",bold=False,sz=10,color=None,bg=None,align="left",
          bdr=False,italic=False,ul=False,wrap=False):
        if isinstance(value, str): value = safe_str(value)
        c = ws.cell(row=row,column=col,value=value)
        kw={"bold":bold,"size":sz,"italic":italic}
        if ul: kw["underline"]="single"
        if color: kw["color"]=color
        c.font=Font(**kw)
        if bg: c.fill=PatternFill("solid",start_color=bg)
        c.alignment=Alignment(horizontal=align,vertical="center",wrap_text=wrap)
        if bdr: c.border=border_all()
        return c

    def merge(r1,c1,r2,c2):
        try:
            ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)
        except:
            pass

    def safe_str(val):
        if val is None: return ""
        s = str(val)
        s = s.replace('\u2013', '-').replace('\u2014', '-')
        s = s.replace('\u2018', "'").replace('\u2019', "'")
        s = s.replace('\u201c', '"').replace('\u201d', '"')
        s = s.replace('\u2022', '*').replace('\u2023', '*')
        s = s.replace('\u2026', '...').replace('\u00a0', ' ')
        s = s.replace('\u2212', '-').replace('\u00d7', 'x')
        result = []
        for ch in s:
            cp = ord(ch)
            if cp < 0x20 and cp not in (0x09, 0x0a, 0x0d): continue
            if 0xD800 <= cp <= 0xDFFF: continue
            if 0xFFFE <= cp <= 0xFFFF: continue
            result.append(ch)
        return ''.join(result)

    for col,wd in {1:3,2:34,3:20,4:16,5:16,6:14,7:14,8:40}.items():
        ws.column_dimensions[get_column_letter(col)].width=wd

    row=1
    ws.row_dimensions[row].height=6; row+=1

    # Header row
    ws.row_dimensions[row].height=26
    w(row,3,"Amounts ($) / No.",bold=True,sz=9,align="center",bg=GRAY)
    w(row,4,"frequency",sz=9,align="center",bg=GRAY,italic=True)
    merge(row,5,row,5)
    c=ws.cell(row=row,column=5,value="APPROVED")
    c.font=Font(bold=True,size=10,color=GREEN_FG)
    c.fill=PatternFill("solid",start_color=GREEN_BG)
    c.alignment=Alignment(horizontal="center",vertical="center")
    c.border=border_all()
    ac = ws.cell(row=row,column=6,value="")
    ac.font=Font(bold=True,size=12,color=GREEN_FG)
    ac.fill=PatternFill("solid",start_color=GREEN_BG)
    ac.alignment=Alignment(horizontal="center",vertical="center")
    ac.border=border_all()
    merge(row,7,row+1,8)
    c=ws.cell(row=row,column=7,value="Update Sheet\nTab Color")
    c.font=Font(bold=True,size=11)
    c.fill=PatternFill("solid",start_color=PURPLE_BG)
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    row+=1

    ws.row_dimensions[row].height=22
    c=ws.cell(row=row,column=5,value="DECLINED")
    c.font=Font(bold=True,size=10,color=RED_FG)
    c.fill=PatternFill("solid",start_color=RED_BG)
    c.alignment=Alignment(horizontal="center",vertical="center")
    c.border=border_all()
    dc = ws.cell(row=row,column=6,value="")
    dc.font=Font(bold=True,size=12,color=RED_FG)
    dc.fill=PatternFill("solid",start_color=RED_BG)
    dc.alignment=Alignment(horizontal="center",vertical="center")
    dc.border=border_all()
    row+=1

    ws.row_dimensions[row].height=6; row+=1
    ws.row_dimensions[row].height=18
    w(row,5,"☐",sz=16,align="center"); row+=1

    # Company name
    ws.row_dimensions[row].height=24
    merge(row,2,row,6)
    c=ws.cell(row=row,column=2,value=data.get("company_name","COMPANY NAME").upper())
    c.font=Font(bold=True,size=13,color=DARK_GOLD)
    c.fill=PatternFill("solid",start_color=LIGHT_YELLOW)
    c.alignment=Alignment(horizontal="left",vertical="center")
    row+=1

    # Offer/Decline
    ws.row_dimensions[row].height=20
    w(row,2,"OFFER / DECLINE",bold=True,ul=True,sz=10)
    w(row,3,"$0.00",sz=10,color=BLUE)
    w(row,4,"daily",sz=9,italic=True)
    w(row,5,"No. of Bank Accts",bold=True,sz=9)
    w(row,6,data.get("num_bank_accounts",1),bold=True,sz=11,align="center")
    row+=1

    for note in data.get("account_notes",[])[:4]:
        ws.row_dimensions[row].height=15
        clr="CC0000" if any(x in note.lower() for x in ["1,000","negative","nsf","returned","overdraft"]) else "000000"
        w(row,2,"*"+note,sz=9,italic=True,color=clr); row+=1
    for _ in range(max(0,3-len(data.get("account_notes",[])))):
        ws.row_dimensions[row].height=14; row+=1

    hb=data.get("holdback_pct",0)
    lv=data.get("leverage_pct",0)
    ws.row_dimensions[row].height=18
    w(row,3,"Holdback %",bold=True,sz=10,align="right")
    w(row,4,"{:.2f}%".format(hb),sz=10,align="center"); row+=1

    ws.row_dimensions[row].height=18
    w(row,2,"SOS",bold=True,ul=True,sz=10)
    w(row,3,"New Holdback %",bold=True,sz=10,align="right")
    w(row,4,"{:.2f}%".format(hb),sz=10,align="center"); row+=1

    ws.row_dimensions[row].height=18
    w(row,2,"Leverage %",bold=True,sz=10,color="CC0000" if lv > 50 else "000000")
    w(row,3,"{:.2f}%".format(lv),bold=True,sz=11,align="center",
      color="CC0000" if lv > 50 else GREEN_FG); row+=1

    ws.row_dimensions[row].height=16
    sos=data.get("sos_info","")
    w(row,2,sos if sos else "Active MM/DD/YYYY",sz=9,italic=True); row+=2

    ws.row_dimensions[row].height=18
    w(row,2,"Court Search",bold=True,ul=True,sz=10); row+=1
    ws.row_dimensions[row].height=16
    court=data.get("court_search_notes","")
    w(row,2,court if court else "*No court records found",sz=9,italic=True,wrap=True); row+=2

    ws.row_dimensions[row].height=20
    acct=data.get("account_number_last4","")
    merge(row,2,row,4)
    c=ws.cell(row=row,column=2,value=acct if acct else "ACCOUNT DETAILS")
    c.font=Font(bold=True,size=12,color=DARK_GOLD)
    c.fill=PatternFill("solid",start_color=GOLD_PALE)
    c.alignment=Alignment(horizontal="left",vertical="center")
    row+=2

    # Current Positions
    ws.row_dimensions[row].height=18
    w(row,2,"Current Positions:",bold=True,ul=True,sz=10)
    total=data.get("total_current_positions",0)
    w(row,3,"${:,.2f}".format(total) if total else "$0.00",bold=True,sz=10,color=BLUE,ul=True)
    w(row,5,"(monthly total)",sz=8,italic=True,color="808080"); row+=1

    for pos in data.get("current_positions",[]):
        ws.row_dimensions[row].height=15
        lender=pos.get("lender",""); amt=pos.get("amount",0)
        freq=pos.get("frequency","weekly"); notes=pos.get("notes","")
        monthly=pos.get("monthly_amount", calc_monthly(amt, freq))
        w(row,2,lender,sz=9,color=BLUE)
        w(row,3,"${:,.2f}".format(amt) if amt else "",sz=9,align="right")
        if freq: w(row,4,"*"+freq,sz=9,italic=True)
        w(row,5,"= ${:,.2f}/mo".format(monthly),sz=8,italic=True,color="808080")
        if notes: w(row,6,"*"+notes,sz=9,italic=True,color="CC0000")
        row+=1

    ws.row_dimensions[row].height=16
    w(row,2,"Other Loans / Positions:",bold=True,sz=10); row+=2

    # Monthly sections
    def month_sort_key(m):
        label = m.get('month_label', '')
        month_map = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
                     'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
        try:
            parts = label.split('-')
            mon = month_map.get(parts[0], 0)
            yr = int(parts[1]) if len(parts) > 1 else 0
            return yr * 100 + mon
        except:
            return 0

    months_sorted = sorted(data.get("months",[]), key=month_sort_key, reverse=True)
    for m in months_sorted:
        label=m.get("month_label",""); period=m.get("period",""); is_mtd=m.get("is_mtd",False)

        ws.row_dimensions[row].height=22
        merge(row,2,row,7)
        hdr="{} (MTD) From {}".format(label,period) if is_mtd and period else label
        c=ws.cell(row=row,column=2,value=hdr)
        c.font=Font(bold=True,size=11)
        c.fill=PatternFill("solid",start_color=MONTH_BG)
        c.alignment=Alignment(horizontal="left",vertical="center")
        row+=1

        ws.row_dimensions[row].height=16
        td=m.get("total_deposits",0)
        td_conf=float(m.get("total_deposits_confidence",1.0))
        w(row,2,"Total deposits:",sz=10)
        w(row,3,"${:,.2f}".format(td),sz=10,color=BLUE)
        if is_mtd: w(row,4,"*calculated *",sz=9,italic=True,color="808080")
        if td_conf < 1.0:
            w(row,7,"⚑ {:.0f}% conf".format(td_conf*100),sz=8,italic=True,color=REVIEW_HEADER,bg=REVIEW_BG)
        row+=1

        ws.row_dimensions[row].height=16
        trd=m.get("true_deposits",0)
        trd_conf=float(m.get("true_deposits_confidence",1.0))
        lbl="True deposits (MTD):" if is_mtd else "True deposits:"
        w(row,2,lbl,sz=10)
        w(row,3,"${:,.2f}".format(trd),sz=10,color=BLUE)
        ntx=m.get("num_transactions",0)
        if is_mtd and ntx: w(row,4,str(ntx),sz=10,align="center",color=BLUE)
        excl=m.get("true_deposit_exclusions","")
        if excl: w(row,5,"*excl. "+excl,sz=8,italic=True,color="808080",wrap=True)
        if trd_conf < 1.0:
            reason = m.get("true_deposits_confidence_reason","")
            w(row,7,"⚑ {:.0f}% conf{}".format(trd_conf*100, ": "+reason if reason else ""),
              sz=8,italic=True,color=REVIEW_HEADER,bg=REVIEW_BG)
        row+=1

        neg=m.get("neg_days",0); nsf=m.get("nsf_count",0); od=m.get("od_count",0)
        bar_label="Neg days # {} / NSF # {} / OD # {}".format(neg,nsf,od)
        bar_bg=NEG_BG if (neg>0 or nsf>0 or od>0) else OK_BG
        bar_fg=RED_FG if (neg>0 or nsf>0 or od>0) else GREEN_FG
        merge(row,2,row,5)
        c=ws.cell(row=row,column=2,value=bar_label)
        c.font=Font(bold=True,size=9,color=bar_fg)
        c.fill=PatternFill("solid",start_color=bar_bg)
        c.alignment=Alignment(horizontal="left",vertical="center")
        nsf_conf=float(m.get("nsf_confidence",1.0))
        if nsf_conf < 1.0:
            w(row,7,"⚑ NSF/OD {:.0f}% conf".format(nsf_conf*100),sz=8,italic=True,color=REVIEW_HEADER,bg=REVIEW_BG)
        row+=1

        adb=m.get("adb",0)
        adb_conf=float(m.get("adb_confidence",1.0))
        w(row,2,"ADB (average daily balance)",sz=10)
        w(row,3,"${:,.2f}".format(adb),sz=10,color=BLUE)
        w(row,4,"*given" if adb_conf==1.0 else "*calculated",sz=9,italic=True,color="808080")
        if adb_conf < 1.0:
            w(row,7,"⚑ {:.0f}% conf".format(adb_conf*100),sz=8,italic=True,color=REVIEW_HEADER,bg=REVIEW_BG)
        row+=1

        dl=m.get("days_below_1000",0)
        dl_conf=float(m.get("days_below_1000_confidence",1.0))
        w(row,2,"Days below $1,000:",sz=10)
        w(row,3,str(dl),sz=10,align="center")
        if dl_conf < 1.0:
            reason = m.get("days_below_1000_confidence_reason","")
            w(row,7,"⚑ {:.0f}% conf{}".format(dl_conf*100, ": "+reason if reason else ""),
              sz=8,italic=True,color=REVIEW_HEADER,bg=REVIEW_BG)
        row+=1

        for fe in m.get("funding_events",[]):
            w(row,2,"*Funded by "+fe.get("funder",""),sz=9,italic=True,color=BLUE)
            amt_fe=fe.get("amount",0)
            w(row,3,"with an amount of ${:,.2f}".format(amt_fe) if amt_fe else "",sz=9,italic=True)
            dt=fe.get("date","")
            if dt: w(row,4,"on "+dt,sz=9,italic=True)
            row+=1

        mnotes=m.get("notes","")
        if mnotes:
            w(row,2,"*"+mnotes,sz=9,italic=True,color="808080",wrap=True); row+=1

        row+=2

    # ─── REVIEW REQUIRED SECTION ──────────────────────────────────────────────
    review_flags = data.get("review_flags", [])
    recon = data.get("balance_reconciliation", {})

    # Always render this section
    ws.row_dimensions[row].height=6; row+=1

    # Section header
    ws.row_dimensions[row].height=26
    merge(row,2,row,8)
    c=ws.cell(row=row,column=2,value="⚑  REVIEW REQUIRED — Items Flagged for Manual Verification")
    c.font=Font(bold=True,size=12,color="FFFFFF")
    c.fill=PatternFill("solid",start_color=REVIEW_HEADER)
    c.alignment=Alignment(horizontal="left",vertical="center")
    row+=1

    # Balance reconciliation box
    ws.row_dimensions[row].height=20
    merge(row,2,row,8)
    c=ws.cell(row=row,column=2,value="BALANCE RECONCILIATION CHECK")
    c.font=Font(bold=True,size=10,color="1F4E79")
    c.fill=PatternFill("solid",start_color=RECON_BG)
    c.alignment=Alignment(horizontal="left",vertical="center")
    row+=1

    if recon:
        reconciles = recon.get("reconciles", True)
        disc = float(recon.get("discrepancy", 0))
        status_text = "PASS — Balances reconcile" if reconciles else "FAIL — Discrepancy: ${:,.2f}".format(abs(disc))
        status_bg = OK_BG if reconciles else NEG_BG
        status_fg = GREEN_FG if reconciles else RED_FG

        merge(row,2,row,5)
        c=ws.cell(row=row,column=2,value=status_text)
        c.font=Font(bold=True,size=10,color=status_fg)
        c.fill=PatternFill("solid",start_color=status_bg)
        c.alignment=Alignment(horizontal="left",vertical="center")
        c.border=border_all()

        w(row,6,"Opening: ${:,.2f}".format(float(recon.get("opening_balance",0))),sz=9,italic=True)
        w(row,7,"Closing: ${:,.2f}".format(float(recon.get("closing_balance",0))),sz=9,italic=True)
        row+=1
    else:
        w(row,2,"Balance reconciliation data not available",sz=9,italic=True,color="808080"); row+=1

    row+=1

    # Column headers for review flags table
    if review_flags:
        ws.row_dimensions[row].height=18
        for col, hdr_text, wd in [
            (2,"Month",8),(3,"Field",22),(4,"Extracted Value",20),
            (5,"Source / Severity",14),(6,"Reason / Action Required",40)
        ]:
            c=ws.cell(row=row,column=col,value=hdr_text)
            c.font=Font(bold=True,size=9,color="FFFFFF")
            c.fill=PatternFill("solid",start_color="595959")
            c.alignment=Alignment(horizontal="center",vertical="center")
            c.border=border_all()
        row+=1

        # Track when we transition from verify flags to confidence flags
        # so we can insert a visual separator
        in_verify_section = True
        for flag in review_flags:
            ftype = flag.get("type","")
            is_verify = ftype.startswith("VERIFY_")

            # Insert separator row when transitioning from verify to confidence flags
            if in_verify_section and not is_verify:
                in_verify_section = False
                ws.row_dimensions[row].height=16
                merge(row,2,row,6)
                c=ws.cell(row=row,column=2,value="— FIRST-PASS CONFIDENCE FLAGS (fields below 100% confidence / unclassified transactions) —")
                c.font=Font(bold=True,size=8,italic=True,color="595959")
                c.fill=PatternFill("solid",start_color="EFEFEF")
                c.alignment=Alignment(horizontal="center",vertical="center")
                row+=1
            ftype = flag.get("type","")

            # Color coding by flag type
            if ftype == "VERIFY_FAILED":
                row_bg = "FF0000"; txt_color = "FFFFFF"
            elif ftype == "VERIFY_PASSED":
                row_bg = GREEN_BG; txt_color = GREEN_FG
            elif ftype == "VERIFY_HIGH":
                row_bg = NEG_BG; txt_color = RED_FG
            elif ftype == "VERIFY_MEDIUM":
                row_bg = "FFE0CC"; txt_color = "CC4400"
            elif ftype == "VERIFY_LOW":
                row_bg = REVIEW_BG; txt_color = REVIEW_HEADER
            elif ftype == "VERIFY_ERROR":
                row_bg = NEG_BG; txt_color = RED_FG
            elif ftype == "UNCLASSIFIED_TXN":
                row_bg = UNCLASS_BG; txt_color = "CC0000"
            elif ftype == "BALANCE_MISMATCH":
                row_bg = NEG_BG; txt_color = RED_FG
            else:
                row_bg = REVIEW_BG; txt_color = REVIEW_HEADER

            conf = flag.get("confidence", 1.0)
            if ftype == "VERIFY_PASSED":
                conf_str = "PASS ✓"
            elif ftype == "VERIFY_FAILED":
                conf_str = "FAIL ✗"
            elif ftype in ("UNCLASSIFIED_TXN", "BALANCE_MISMATCH", "VERIFY_ERROR"):
                conf_str = ftype.replace("_", " ")
            elif ftype in ("VERIFY_HIGH","VERIFY_MEDIUM","VERIFY_LOW"):
                severity = ftype.replace("VERIFY_","")
                conf_str = "2nd PASS — {}".format(severity)
            else:
                conf_str = "{:.0f}%".format(float(conf)*100)

            ws.row_dimensions[row].height=30
            w(row,2,flag.get("month","—"),sz=9,bg=row_bg,bdr=True,align="center")
            w(row,3,flag.get("field",""),sz=9,bg=row_bg,bdr=True,wrap=True)
            w(row,4,flag.get("value",""),sz=9,bg=row_bg,bdr=True,wrap=True)
            c=ws.cell(row=row,column=5,value=conf_str)
            c.font=Font(bold=True,size=9,color=txt_color)
            c.fill=PatternFill("solid",start_color=row_bg)
            c.alignment=Alignment(horizontal="center",vertical="center")
            c.border=border_all()
            w(row,6,flag.get("reason",""),sz=9,bg=row_bg,bdr=True,wrap=True)
            row+=1

    else:
        # No flags - green all-clear
        ws.row_dimensions[row].height=22
        merge(row,2,row,8)
        c=ws.cell(row=row,column=2,value="✓  No review flags — All fields extracted at 100% confidence, all transactions classified")
        c.font=Font(bold=True,size=10,color=GREEN_FG)
        c.fill=PatternFill("solid",start_color=GREEN_BG)
        c.alignment=Alignment(horizontal="left",vertical="center")
        c.border=border_all()
        row+=1

    row+=1

    ws.freeze_panes="B7"
    wb.active.sheet_properties.tabColor="C8962A"
    out=io.BytesIO(); wb.save(out); out.seek(0)
    return out

@app.route('/login', methods=['GET','POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username','').strip()
        password = request.form.get('password','').strip()
        if USERS.get(username) == password:
            session['logged_in'] = True
            session['username'] = username
            return redirect(url_for('index'))
        error = 'Invalid username or password'
    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    history = load_history()
    return render_template('index.html', history=history)

@app.route('/analyze', methods=['POST'])
@login_required
def analyze():
    if 'files' not in request.files:
        return jsonify({"error": "No files uploaded"}), 400
    files = request.files.getlist('files')
    company_name = request.form.get('company_name', '')
    entry_id = request.form.get('entry_id', '')
    if not files or all(f.filename == '' for f in files):
        return jsonify({"error": "No files selected"}), 400

    combined_text = ""
    for file in files:
        if file and allowed_file(file.filename):
            fname = secure_filename(file.filename)
            fpath = os.path.join(app.config['UPLOAD_FOLDER'], fname)
            file.save(fpath)
            try:
                combined_text += "\n\n=== FILE: {} ===\n".format(fname) + extract_text(fpath)
            except Exception as e:
                return jsonify({"error": "Failed to read {}: {}".format(fname, str(e))}), 500
            finally:
                if os.path.exists(fpath):
                    os.remove(fpath)
        else:
            return jsonify({"error": "Unsupported file: {}".format(file.filename)}), 400

    if not combined_text.strip():
        return jsonify({"error": "No text extracted"}), 400

    # Load existing entry now (in the request thread) before handing off
    existing_entry = None
    if entry_id:
        existing_entry = load_entry(entry_id)

    # Create job and kick off background thread
    job_id = str(uuid.uuid4())
    job_set(job_id, status="running", progress=5, message="Starting analysis...")

    t = threading.Thread(
        target=run_analysis_job,
        args=(job_id, combined_text, company_name, entry_id, existing_entry),
        daemon=True
    )
    t.start()

    return jsonify({"job_id": job_id}), 202


@app.route('/analyze/status/<job_id>')
@login_required
def analyze_status(job_id):
    job = job_get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "status":   job.get("status", "unknown"),
        "progress": job.get("progress", 0),
        "message":  job.get("message", ""),
        "error":    job.get("error", ""),
        "filename": job.get("filename", "")
    })


@app.route('/analyze/download/<job_id>')
@login_required
def analyze_download(job_id):
    job = job_get(job_id)
    if not job or job.get("status") != "done":
        return "Job not ready or not found", 404
    excel_bytes = job.get("excel_bytes")
    filename = job.get("filename", "analysis.xlsx")
    # Clean up job from memory after download
    with JOBS_LOCK:
        JOBS.pop(job_id, None)
    return send_file(
        io.BytesIO(excel_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/history/<entry_id>/download')
@login_required
def download_history(entry_id):
    entry = load_entry(entry_id)
    if not entry: return "Not found", 404
    try:
        clean_data = sanitize_data(entry['data'])
        excel = build_excel(clean_data)
        excel_bytes = excel.read()
    except:
        excel_bytes = entry['excel']
    safe=re.sub(r'[^\w\s-]','',entry['company_name']).strip().replace(' ','_')
    return send_file(io.BytesIO(excel_bytes),as_attachment=True,
                     download_name=safe+"_analysis.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/history/<entry_id>/delete', methods=['POST'])
@login_required
def delete_history(entry_id):
    path = os.path.join(app.config['HISTORY_FOLDER'], entry_id + '.pkl')
    if os.path.exists(path): os.remove(path)
    return redirect(url_for('index'))

if __name__=='__main__':
    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT', 5001)))
