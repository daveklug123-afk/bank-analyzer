import os, json, re, io, csv, anthropic
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename
import pdfplumber, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.environ.get('UPLOAD_FOLDER', '/tmp/uploads')
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024
ALLOWED_EXTENSIONS = {'pdf', 'csv', 'txt'}

def allowed_file(f): return '.' in f and f.rsplit('.',1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_pdf(path):
    text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t: text += t + "\n"
    return text

def extract_text(path):
    ext = path.rsplit('.',1)[1].lower()
    if ext == 'pdf': return extract_text_from_pdf(path)
    with open(path,'r',encoding='utf-8',errors='ignore') as f: return f.read()

def parse_with_claude(raw_text, company_name=""):
    client = anthropic.Anthropic()
    prompt = f"""You are an expert MCA (Merchant Cash Advance) underwriter analyzing Chase bank statements.

IMPORTANT CONTEXT - You are the LENDER evaluating whether to lend money to this business.

Analyze the bank statement text below and extract ALL data. Return ONLY valid JSON, no markdown.

Bank Statement Text:
{raw_text[:25000]}

EXTRACTION RULES:
1. true_deposits = total_deposits MINUS:
   - Fedwire/incoming wire credits that are MCA fundings (look for lender names in B/O field like "SQ ADVANCE LLC", "GARDEN FUNDING", "KYLE CAPITAL LLC", "NYC ADVANCE GROUP", "EMMY CAPITAL", "PARKVIEW ADVANCE", "SHILENO LLC" used as passthrough)
   - Online transfers FROM other own accounts (e.g. "Online Transfer From Chk ...6837")
   - Book Transfer Credits between own accounts
   - Hunter Caroline reversal credits (these are reversals of MCA debits, not real revenue)
2. current_positions = ALL active MCA lenders found as recurring ACH debits in electronic withdrawals:
   - Look for: Hunter Caroline, LG Funding LLC, Garden Funding L, SQ Advance, Libertasfunding, Catalystadvance, Kyle Capital, EMMY CAPITAL GROUP, Parkview Advance, NYC Advance Group
   - Use the most recent repayment amount for each
3. adb = average of all daily ending balances shown in "DAILY ENDING BALANCE" section
4. days_below_1000 = count days where daily ending balance < 1000 (including negative days)
5. neg_days = count days where daily ending balance is negative (< 0)
6. nsf_count = count "NSF" or "Returned Item" or "Insufficient" entries in fees
7. od_count = count overdraft fees
8. funding_events = Fedwire/incoming wire credits that are MCA fundings - identify funder from B/O field
9. For "Shileno LLC" wires - these appear to be client payments (fence/construction work), NOT MCA fundings - INCLUDE in true deposits
10. LIBERTAS FUNDING #1 stopped after 03/09 per the analysis sheet - note this
11. holdback_pct: estimate based on total MCA repayments / total deposits (as percentage)

Return this exact JSON:
{{
  "company_name": "CHIEF TOP REMODELING INC",
  "account_number_last4": "9526",
  "num_bank_accounts": 1,
  "offer_decline": "DECLINE",
  "holdback_pct": 13.78,
  "sos_info": "Active 07/21/2022",
  "court_search_notes": "",
  "account_notes": ["inconsistent revenue", "many days < $1,000 in Feb"],
  "total_current_positions": 95405.04,
  "current_positions": [
    {{"lender": "NYC ADVANCE GROUP", "amount": 0, "frequency": "weekly", "notes": "re-payments not started yet / not found"}},
    {{"lender": "Parkview Advance", "amount": 2500.00, "frequency": "weekly", "notes": "only 1 re-payment on 04/21"}},
    {{"lender": "EMMY CAPITAL GROUP", "amount": 3295.46, "frequency": "weekly", "notes": "only 1 re-payment on 04/24"}},
    {{"lender": "Kyle Capital", "amount": 3182.00, "frequency": "weekly", "notes": ""}},
    {{"lender": "LIBERTAS FUNDING #2", "amount": 803.55, "frequency": "weekly", "notes": ""}},
    {{"lender": "LG Funding LLC #2", "amount": 1995.00, "frequency": "weekly", "notes": ""}},
    {{"lender": "LG Funding LLC #1", "amount": 1750.00, "frequency": "weekly", "notes": ""}},
    {{"lender": "SQ ADVANCE", "amount": 2900.00, "frequency": "weekly", "notes": ""}},
    {{"lender": "Catalystadvance", "amount": 1421.00, "frequency": "weekly", "notes": ""}},
    {{"lender": "GARDEN FUNDING", "amount": 2998.00, "frequency": "weekly", "notes": "$1,986.18 weekly before 03/06"}},
    {{"lender": "HUNTER CAROLINE", "amount": 3006.25, "frequency": "weekly", "notes": "seems an MCA"}},
    {{"lender": "LIBERTAS FUNDING #1", "amount": 1060.70, "frequency": "weekly", "notes": "stopped after 03/09, last repayment $1,061.30 on 03/09"}}
  ],
  "months": [
    {{
      "month_label": "Apr-26",
      "period": "04/01 to 04/24",
      "is_mtd": true,
      "total_deposits": 396900.00,
      "true_deposits": 282850.00,
      "true_deposit_notes": "incl. Fedwire Credit B/O: Shileno LLC Bnf=Chief Top Remodeling amounting $550",
      "neg_days": 1,
      "nsf_count": 0,
      "od_count": 0,
      "num_transactions": 24,
      "adb": 10519.22,
      "days_below_1000": 7,
      "funding_events": [
        {{"funder": "NYC Group Advance", "amount": 36000.00, "date": "04/23"}},
        {{"funder": "PARKVIEW ADVANCE", "amount": 29100.00, "date": "04/14"}},
        {{"funder": "EMMY CAPITAL", "amount": 47500.00, "date": "04/15"}}
      ],
      "notes": ""
    }},
    {{
      "month_label": "Mar-26",
      "period": "02/28 to 03/31",
      "is_mtd": false,
      "total_deposits": 1181009.65,
      "true_deposits": 1128329.00,
      "true_deposit_notes": "incl. Fedwire Credit B/O: Shileno LLC Bnf=Chief Top Remodeling amounting $2,500",
      "neg_days": 2,
      "nsf_count": 0,
      "od_count": 0,
      "num_transactions": 320,
      "adb": 6620.58,
      "days_below_1000": 8,
      "funding_events": [
        {{"funder": "Kyle Capital LLC", "amount": 47500.00, "date": "03/19"}}
      ],
      "notes": ""
    }},
    {{
      "month_label": "Feb-26",
      "period": "01/31 to 02/27",
      "is_mtd": false,
      "total_deposits": 816358.37,
      "true_deposits": 739964.00,
      "true_deposit_notes": "incl. Fedwire Credit B/O: Shileno LLC Bnf=Chief Top Remodeling amounting $2,000; incl. Book Transfer Credit B/O: ...",
      "neg_days": 0,
      "nsf_count": 0,
      "od_count": 0,
      "num_transactions": 225,
      "adb": 7802.52,
      "days_below_1000": 9,
      "funding_events": [
        {{"funder": "Sq Advance LLC", "amount": 47500.00, "date": "02/12"}},
        {{"funder": "Garden Funding LLC", "amount": 25573.12, "date": "02/27"}}
      ],
      "notes": "1 stop payment fee of $30"
    }}
  ]
}}

Now RE-ANALYZE the actual statement text provided and return accurate JSON based on what you actually find in the text. Use the structure above as a template but populate with REAL extracted values.
Company name override: "{company_name if company_name else 'auto-detect'}"
"""
    msg = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=6000,
        messages=[{"role":"user","content":prompt}]
    )
    raw = msg.content[0].text.strip()
    raw = re.sub(r'^```json\s*','',raw); raw = re.sub(r'^```\s*','',raw); raw = re.sub(r'\s*```$','',raw)
    return json.loads(raw)

def build_excel(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Analysis"

    # Color palette
    GOLD="#C8962A"; GOLD_PALE="FFF8E1"; LIGHT_YELLOW="FFFF99"; DARK_GOLD="8B6914"
    GREEN_BG="C6EFCE"; GREEN_FG="006100"; RED_BG="FFC7CE"; RED_FG="9C0006"
    ORANGE_BG="FFE0B2"; BLUE="0070C0"; PURPLE_BG="EAD5F5"
    NEG_BG="FFC7CE"; OK_BG="C6EFCE"; GRAY="F2F2F2"; MED_GRAY="D9D9D9"
    MONTH_BG="FFD966"  # gold/amber for month headers

    thin = Side(style='thin'); med = Side(style='medium')
    def border(l=False,r=False,t=False,b=False,all=False,med_all=False):
        s = med if med_all else thin
        if all or med_all: return Border(left=s,right=s,top=s,bottom=s)
        return Border(left=thin if l else Side(style=None),right=thin if r else Side(style=None),
                      top=thin if t else Side(style=None),bottom=thin if b else Side(style=None))

    def w(row,col,value="",bold=False,sz=10,color=None,bg=None,align="left",
          bdr=None,italic=False,ul=False,wrap=False,num_fmt=None):
        c = ws.cell(row=row,column=col,value=value)
        kw={"bold":bold,"size":sz,"italic":italic}
        if ul: kw["underline"]="single"
        if color: kw["color"]=color
        c.font=Font(**kw)
        if bg: c.fill=PatternFill("solid",start_color=bg)
        c.alignment=Alignment(horizontal=align,vertical="center",wrap_text=wrap)
        if bdr: c.border=bdr
        if num_fmt: c.number_format=num_fmt
        return c

    def merge(r1,c1,r2,c2): ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)

    # Column widths
    for col,wd in {1:3,2:34,3:20,4:16,5:16,6:14,7:14,8:36}.items():
        ws.column_dimensions[get_column_letter(col)].width=wd

    row=1
    ws.row_dimensions[row].height=6; row+=1

    # ── Row 2: Headers + APPROVED/DECLINED ──
    ws.row_dimensions[row].height=26
    w(row,3,"Amounts ($) / No.",bold=True,sz=9,align="center",bg=GRAY)
    w(row,4,"frequency",sz=9,align="center",bg=GRAY,italic=True)
    # APPROVED
    merge(row,5,row,5)
    c=ws.cell(row=row,column=5,value="APPROVED")
    c.font=Font(bold=True,size=10,color=GREEN_FG)
    c.fill=PatternFill("solid",start_color=GREEN_BG)
    c.alignment=Alignment(horizontal="center",vertical="center")
    c.border=border(all=True)
    w(row,6,"☑" if data.get("offer_decline")=="OFFER" else "☐",sz=14,align="center",bg=GREEN_BG)
    # Tab Color label
    merge(row,7,row+1,8)
    c=ws.cell(row=row,column=7,value="Update Sheet\nTab Color")
    c.font=Font(bold=True,size=11); c.fill=PatternFill("solid",start_color=PURPLE_BG)
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    row+=1

    # Row 3: DECLINED
    ws.row_dimensions[row].height=22
    c=ws.cell(row=row,column=5,value="DECLINED")
    c.font=Font(bold=True,size=10,color=RED_FG)
    c.fill=PatternFill("solid",start_color=RED_BG)
    c.alignment=Alignment(horizontal="center",vertical="center")
    c.border=border(all=True)
    w(row,6,"☑" if data.get("offer_decline")=="DECLINE" else "☐",sz=14,align="center",bg=RED_BG)
    row+=1

    ws.row_dimensions[row].height=6; row+=1   # row 4 spacer
    ws.row_dimensions[row].height=18          # row 5 checkbox row
    w(row,5,"☐",sz=16,align="center")
    row+=1

    # ── Row 6: Company Name ──
    ws.row_dimensions[row].height=24
    merge(row,2,row,6)
    c=ws.cell(row=row,column=2,value=data.get("company_name","COMPANY NAME").upper())
    c.font=Font(bold=True,size=13,color=DARK_GOLD)
    c.fill=PatternFill("solid",start_color=LIGHT_YELLOW)
    c.alignment=Alignment(horizontal="left",vertical="center")
    row+=1

    # ── OFFER / DECLINE line ──
    ws.row_dimensions[row].height=20
    w(row,2,"OFFER / DECLINE",bold=True,ul=True,sz=10)
    w(row,3,"$0.00",sz=10,color=BLUE)
    w(row,4,"daily",sz=9,italic=True)
    w(row,5,"No. of Bank Accts",bold=True,sz=9)
    w(row,6,data.get("num_bank_accounts",1),bold=True,sz=11,align="center")
    row+=1

    # Account notes
    for note in data.get("account_notes",[])[:4]:
        ws.row_dimensions[row].height=15
        clr="CC0000" if any(x in note.lower() for x in ["1,000","negative","nsf"]) else "000000"
        w(row,2,f"*{note}",sz=9,italic=True,color=clr)
        row+=1

    # Pad to at least 3 note rows
    for _ in range(max(0,3-len(data.get("account_notes",[])))):
        ws.row_dimensions[row].height=14; row+=1

    # ── Holdback ──
    ws.row_dimensions[row].height=18
    w(row,3,"Holdback %",bold=True,sz=10,align="right")
    hb=data.get("holdback_pct",0)
    w(row,4,f"{hb:.2f}%",sz=10,align="center")
    row+=1

    ws.row_dimensions[row].height=18
    w(row,2,"SOS",bold=True,ul=True,sz=10)
    w(row,3,"New Holdback %",bold=True,sz=10,align="right")
    w(row,4,f"{hb:.2f}%",sz=10,align="center")
    row+=1

    ws.row_dimensions[row].height=16
    sos=data.get("sos_info","")
    w(row,2,sos if sos else "Active MM/DD/YYYY",sz=9,italic=True)
    row+=2

    # ── Court Search ──
    ws.row_dimensions[row].height=18
    w(row,2,"Court Search",bold=True,ul=True,sz=10); row+=1
    ws.row_dimensions[row].height=16
    court=data.get("court_search_notes","")
    w(row,2,court if court else "*No court records found",sz=9,italic=True,wrap=True); row+=2

    # ── Account # ──
    ws.row_dimensions[row].height=20
    acct=data.get("account_number_last4","")
    merge(row,2,row,4)
    c=ws.cell(row=row,column=2,value=f"{acct}" if acct else "ACCOUNT DETAILS")
    c.font=Font(bold=True,size=12,color=DARK_GOLD)
    c.fill=PatternFill("solid",start_color=GOLD_PALE)
    c.alignment=Alignment(horizontal="left",vertical="center")
    row+=2

    # ── Current Positions ──
    ws.row_dimensions[row].height=18
    w(row,2,"Current Positions:",bold=True,ul=True,sz=10)
    total=data.get("total_current_positions",0)
    w(row,3,f"${total:,.2f}" if total else "$0.00",bold=True,sz=10,color=BLUE,ul=True)
    row+=1

    for pos in data.get("current_positions",[]):
        ws.row_dimensions[row].height=15
        lender=pos.get("lender",""); amt=pos.get("amount",0)
        freq=pos.get("frequency","weekly"); notes=pos.get("notes","")
        lclr=BLUE if amt and amt>0 else "808080"
        w(row,2,lender,sz=9,color=lclr)
        if amt: w(row,3,f"${amt:,.2f}",sz=9,align="right")
        if freq: w(row,4,f"*{freq}" if freq else "",sz=9,italic=True)
        if notes: w(row,5,f"*{notes}",sz=9,italic=True,color="CC0000")
        row+=1

    ws.row_dimensions[row].height=16
    w(row,2,"Other Loans / Positions:",bold=True,sz=10); row+=2

    # ── Monthly Sections ──
    months=data.get("months",[])
    for i,m in enumerate(months):
        label=m.get("month_label","")
        period=m.get("period","")
        is_mtd=m.get("is_mtd",False)

        # Month header
        ws.row_dimensions[row].height=22
        merge(row,2,row,7)
        hdr=f"{label} (MTD) From {period}" if is_mtd and period else label
        c=ws.cell(row=row,column=2,value=hdr)
        c.font=Font(bold=True,size=11)
        c.fill=PatternFill("solid",start_color=MONTH_BG)
        c.alignment=Alignment(horizontal="left",vertical="center")
        row+=1

        # Total deposits
        ws.row_dimensions[row].height=16
        td=m.get("total_deposits",0)
        w(row,2,"Total deposits:",sz=10)
        w(row,3,f"${td:,.2f}",sz=10,color=BLUE)
        if is_mtd: w(row,4,"*calculated *",sz=9,italic=True,color="808080")
        row+=1

        # True deposits
        ws.row_dimensions[row].height=16
        trd=m.get("true_deposits",0)
        lbl="True deposits (MTD):" if is_mtd else "True deposits:"
        w(row,2,lbl,sz=10)
        w(row,3,f"${trd:,.2f}",sz=10,color=BLUE)
        ntx=m.get("num_transactions",0)
        if is_mtd and ntx:
            w(row,4,str(ntx),sz=10,align="center",color=BLUE)
        tnote=m.get("true_deposit_notes","")
        if tnote: w(row,5,f"*incl. {tnote}",sz=8,italic=True,color="808080",wrap=True)
        row+=1

        # Neg/NSF/OD bar
        ws.row_dimensions[row].height=16
        neg=m.get("neg_days",0); nsf=m.get("nsf_count",0); od=m.get("od_count",0)
        bar_label=f"Neg days # {neg} / NSF # {nsf} / OD # {od}"
        bar_bg=NEG_BG if (neg>0 or nsf>0 or od>0) else OK_BG
        bar_fg=RED_FG if (neg>0 or nsf>0 or od>0) else GREEN_FG
        merge(row,2,row,5)
        c=ws.cell(row=row,column=2,value=bar_label)
        c.font=Font(bold=True,size=9,color=bar_fg)
        c.fill=PatternFill("solid",start_color=bar_bg)
        c.alignment=Alignment(horizontal="left",vertical="center")
        row+=1

        # ADB
        ws.row_dimensions[row].height=16
        adb=m.get("adb",0)
        w(row,2,"ADB (average daily balance)",sz=10)
        w(row,3,f"${adb:,.2f}",sz=10,color=BLUE)
        w(row,4,"*calculated",sz=9,italic=True,color="808080")
        row+=1

        # Days below 1000
        ws.row_dimensions[row].height=16
        dl=m.get("days_below_1000",0)
        w(row,2,"Days below $1,000:",sz=10)
        w(row,3,str(dl),sz=10,align="center")
        row+=1

        # Funding events
        for fe in m.get("funding_events",[]):
            ws.row_dimensions[row].height=15
            w(row,2,f"*Funded by {fe.get('funder','')}",sz=9,italic=True,color=BLUE)
            amt_fe=fe.get('amount',0)
            w(row,3,f"with an amount of ${amt_fe:,.2f}" if amt_fe else "",sz=9,italic=True)
            dt=fe.get('date','')
            if dt: w(row,4,f"on {dt}",sz=9,italic=True)
            row+=1

        mnotes=m.get("notes","")
        if mnotes:
            ws.row_dimensions[row].height=15
            w(row,2,f"*{mnotes}",sz=9,italic=True,color="808080",wrap=True); row+=1

        row+=2  # spacing

    ws.freeze_panes="B7"
    wb.active.sheet_properties.tabColor="C8962A"
    out=io.BytesIO(); wb.save(out); out.seek(0)
    return out

@app.route('/')
def index(): return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    if 'files' not in request.files: return jsonify({"error":"No files uploaded"}),400
    files=request.files.getlist('files')
    company_name=request.form.get('company_name','')
    if not files or all(f.filename=='' for f in files): return jsonify({"error":"No files selected"}),400

    combined_text=""
    for file in files:
        if file and allowed_file(file.filename):
            fname=secure_filename(file.filename)
            fpath=os.path.join(app.config['UPLOAD_FOLDER'],fname)
            file.save(fpath)
            try:
                combined_text+=f"\n\n=== FILE: {fname} ===\n{extract_text(fpath)}"
            except Exception as e:
                return jsonify({"error":f"Failed to read {fname}: {str(e)}"}),500
            finally:
                if os.path.exists(fpath): os.remove(fpath)
        else: return jsonify({"error":f"Unsupported file: {file.filename}"}),400

    if not combined_text.strip(): return jsonify({"error":"No text extracted"}),400

    try: parsed=parse_with_claude(combined_text,company_name)
    except Exception as e: return jsonify({"error":f"AI parsing failed: {str(e)}"}),500

    try: excel=build_excel(parsed)
    except Exception as e: return jsonify({"error":f"Excel generation failed: {str(e)}"}),500

    safe=re.sub(r'[^\w\s-]','',parsed.get("company_name","analysis")).strip().replace(' ','_')
    return send_file(excel,as_attachment=True,download_name=f"{safe}_analysis.xlsx",
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__=='__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT', 5001)))
