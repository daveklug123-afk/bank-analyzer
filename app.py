import os, json, re, io, anthropic
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename
import pdfplumber, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.environ.get('UPLOAD_FOLDER', '/tmp/uploads')
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024
ALLOWED_EXTENSIONS = {'pdf', 'csv', 'txt'}
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

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
    cn = company_name if company_name else 'auto-detect'
    prompt = (
        "You are an expert MCA underwriter analyzing bank statements. You are the LENDER.\n"
        "Return ONLY valid JSON, no markdown, no explanation.\n\n"
        "BANK STATEMENT TEXT (may contain multiple months):\n"
        + raw_text[:80000] +
        "\n\nCRITICAL INSTRUCTIONS:\n"
        "1. Find EVERY statement period. Look for CHECKING SUMMARY headers and date ranges.\n"
        "2. Extract ALL months found - do not skip any.\n"
        "3. Most recent partial month gets is_mtd = true.\n\n"
        "FOR EACH MONTH:\n"
        "- total_deposits: Deposits and Additions total from CHECKING SUMMARY\n"
        "- true_deposits: total_deposits minus MCA funding Fedwires (B/O field shows lender names like SQ ADVANCE, GARDEN FUNDING, KYLE CAPITAL, NYC ADVANCE GROUP, EMMY CAPITAL, PARKVIEW ADVANCE) and minus Online Transfer From Chk entries\n"
        "- Shileno LLC wires are real client payments - KEEP in true deposits\n"
        "- adb: average of all values in DAILY ENDING BALANCE table\n"
        "- neg_days: count of negative balances in DAILY ENDING BALANCE table\n"
        "- days_below_1000: count of balances under 1000 in DAILY ENDING BALANCE table\n"
        "- funding_events: incoming Fedwire credits that are MCA loans\n\n"
        "FOR CURRENT POSITIONS: find all recurring ACH debits in Electronic Withdrawals. Use most recent amount per lender. Do NOT calculate total_current_positions - leave it as 0.\n\n"
        "Return this JSON structure:\n"
        '{"company_name":"string","account_number_last4":"string","num_bank_accounts":1,'
        '"offer_decline":"DECLINE","holdback_pct":0.0,"sos_info":"","court_search_notes":"",'
        '"account_notes":[],"total_current_positions":0.0,'
        '"current_positions":[{"lender":"name","amount":0.0,"frequency":"weekly","notes":""}],'
        '"months":[{"month_label":"Mon-YY","period":"MM/DD to MM/DD","is_mtd":false,'
        '"total_deposits":0.0,"true_deposits":0.0,"true_deposit_notes":"",'
        '"neg_days":0,"nsf_count":0,"od_count":0,"num_transactions":0,'
        '"adb":0.0,"days_below_1000":0,'
        '"funding_events":[{"funder":"name","amount":0.0,"date":"MM/DD"}],'
        '"notes":""}]}\n\n'
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
    return json.loads(raw)

def build_excel(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Analysis"

    GOLD_PALE="FFF8E1"; LIGHT_YELLOW="FFFF99"; DARK_GOLD="8B6914"
    GREEN_BG="C6EFCE"; GREEN_FG="006100"; RED_BG="FFC7CE"; RED_FG="9C0006"
    BLUE="0070C0"; PURPLE_BG="EAD5F5"; GRAY="F2F2F2"
    NEG_BG="FFC7CE"; OK_BG="C6EFCE"; MONTH_BG="FFD966"

    thin = Side(style='thin')
    def border_all():
        return Border(left=thin,right=thin,top=thin,bottom=thin)

    def w(row,col,value="",bold=False,sz=10,color=None,bg=None,align="left",
          bdr=False,italic=False,ul=False,wrap=False):
        c = ws.cell(row=row,column=col,value=value)
        kw={"bold":bold,"size":sz,"italic":italic}
        if ul: kw["underline"]="single"
        if color: kw["color"]=color
        c.font=Font(**kw)
        if bg: c.fill=PatternFill("solid",start_color=bg)
        c.alignment=Alignment(horizontal=align,vertical="center",wrap_text=wrap)
        if bdr: c.border=border_all()
        return c

    def merge(r1,c1,r2,c2): ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)

    for col,wd in {1:3,2:34,3:20,4:16,5:16,6:14,7:14,8:36}.items():
        ws.column_dimensions[get_column_letter(col)].width=wd

    row=1
    ws.row_dimensions[row].height=6; row+=1

    ws.row_dimensions[row].height=26
    w(row,3,"Amounts ($) / No.",bold=True,sz=9,align="center",bg=GRAY)
    w(row,4,"frequency",sz=9,align="center",bg=GRAY,italic=True)
    merge(row,5,row,5)
    c=ws.cell(row=row,column=5,value="APPROVED")
    c.font=Font(bold=True,size=10,color=GREEN_FG)
    c.fill=PatternFill("solid",start_color=GREEN_BG)
    c.alignment=Alignment(horizontal="center",vertical="center")
    c.border=border_all()
    w(row,6,"☑" if data.get("offer_decline")=="OFFER" else "☐",sz=14,align="center",bg=GREEN_BG)
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
    w(row,6,"☑" if data.get("offer_decline")=="DECLINE" else "☐",sz=14,align="center",bg=RED_BG)
    row+=1

    ws.row_dimensions[row].height=6; row+=1
    ws.row_dimensions[row].height=18
    w(row,5,"☐",sz=16,align="center"); row+=1

    ws.row_dimensions[row].height=24
    merge(row,2,row,6)
    c=ws.cell(row=row,column=2,value=data.get("company_name","COMPANY NAME").upper())
    c.font=Font(bold=True,size=13,color=DARK_GOLD)
    c.fill=PatternFill("solid",start_color=LIGHT_YELLOW)
    c.alignment=Alignment(horizontal="left",vertical="center")
    row+=1

    ws.row_dimensions[row].height=20
    w(row,2,"OFFER / DECLINE",bold=True,ul=True,sz=10)
    w(row,3,"$0.00",sz=10,color=BLUE)
    w(row,4,"daily",sz=9,italic=True)
    w(row,5,"No. of Bank Accts",bold=True,sz=9)
    w(row,6,data.get("num_bank_accounts",1),bold=True,sz=11,align="center")
    row+=1

    for note in data.get("account_notes",[])[:4]:
        ws.row_dimensions[row].height=15
        clr="CC0000" if any(x in note.lower() for x in ["1,000","negative","nsf"]) else "000000"
        w(row,2,"*"+note,sz=9,italic=True,color=clr); row+=1

    for _ in range(max(0,3-len(data.get("account_notes",[])))):
        ws.row_dimensions[row].height=14; row+=1

    ws.row_dimensions[row].height=18
    w(row,3,"Holdback %",bold=True,sz=10,align="right")
    hb=data.get("holdback_pct",0)
    w(row,4,"{:.2f}%".format(hb),sz=10,align="center"); row+=1

    ws.row_dimensions[row].height=18
    w(row,2,"SOS",bold=True,ul=True,sz=10)
    w(row,3,"New Holdback %",bold=True,sz=10,align="right")
    w(row,4,"{:.2f}%".format(hb),sz=10,align="center"); row+=1

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

    ws.row_dimensions[row].height=18
    w(row,2,"Current Positions:",bold=True,ul=True,sz=10)
    total=data.get("total_current_positions",0)
    w(row,3,"${:,.2f}".format(total) if total else "$0.00",bold=True,sz=10,color=BLUE,ul=True); row+=1

    for pos in data.get("current_positions",[]):
        ws.row_dimensions[row].height=15
        lender=pos.get("lender",""); amt=pos.get("amount",0)
        bal=pos.get("outstanding_balance",0)
        freq=pos.get("frequency","weekly"); notes=pos.get("notes","")
        lclr=BLUE if amt and amt>0 else "808080"
        w(row,2,lender,sz=9,color=lclr)
        w(row,3,"${:,.2f}".format(bal) if bal else "",sz=9,align="right",color=BLUE)
        if freq: w(row,4,"*"+freq,sz=9,italic=True)
        if notes: w(row,5,"*"+notes,sz=9,italic=True,color="CC0000")
        row+=1

    ws.row_dimensions[row].height=16
    w(row,2,"Other Loans / Positions:",bold=True,sz=10); row+=2

    for i,m in enumerate(data.get("months",[])):
        label=m.get("month_label","")
        period=m.get("period","")
        is_mtd=m.get("is_mtd",False)

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
        w(row,2,"Total deposits:",sz=10)
        w(row,3,"${:,.2f}".format(td),sz=10,color=BLUE)
        if is_mtd: w(row,4,"*calculated *",sz=9,italic=True,color="808080")
        row+=1

        ws.row_dimensions[row].height=16
        trd=m.get("true_deposits",0)
        lbl="True deposits (MTD):" if is_mtd else "True deposits:"
        w(row,2,lbl,sz=10)
        w(row,3,"${:,.2f}".format(trd),sz=10,color=BLUE)
        ntx=m.get("num_transactions",0)
        if is_mtd and ntx: w(row,4,str(ntx),sz=10,align="center",color=BLUE)
        tnote=m.get("true_deposit_notes","")
        if tnote: w(row,5,"*incl. "+tnote,sz=8,italic=True,color="808080",wrap=True)
        row+=1

        ws.row_dimensions[row].height=16
        neg=m.get("neg_days",0); nsf=m.get("nsf_count",0); od=m.get("od_count",0)
        bar_label="Neg days # {} / NSF # {} / OD # {}".format(neg,nsf,od)
        bar_bg=NEG_BG if (neg>0 or nsf>0 or od>0) else OK_BG
        bar_fg=RED_FG if (neg>0 or nsf>0 or od>0) else GREEN_FG
        merge(row,2,row,5)
        c=ws.cell(row=row,column=2,value=bar_label)
        c.font=Font(bold=True,size=9,color=bar_fg)
        c.fill=PatternFill("solid",start_color=bar_bg)
        c.alignment=Alignment(horizontal="left",vertical="center")
        row+=1

        ws.row_dimensions[row].height=16
        adb=m.get("adb",0)
        w(row,2,"ADB (average daily balance)",sz=10)
        w(row,3,"${:,.2f}".format(adb),sz=10,color=BLUE)
        w(row,4,"*calculated",sz=9,italic=True,color="808080"); row+=1

        ws.row_dimensions[row].height=16
        dl=m.get("days_below_1000",0)
        w(row,2,"Days below $1,000:",sz=10)
        w(row,3,str(dl),sz=10,align="center"); row+=1

        for fe in m.get("funding_events",[]):
            ws.row_dimensions[row].height=15
            w(row,2,"*Funded by "+fe.get("funder",""),sz=9,italic=True,color=BLUE)
            amt_fe=fe.get("amount",0)
            w(row,3,"with an amount of ${:,.2f}".format(amt_fe) if amt_fe else "",sz=9,italic=True)
            dt=fe.get("date","")
            if dt: w(row,4,"on "+dt,sz=9,italic=True)
            row+=1

        mnotes=m.get("notes","")
        if mnotes:
            ws.row_dimensions[row].height=15
            w(row,2,"*"+mnotes,sz=9,italic=True,color="808080",wrap=True); row+=1

        row+=2

    ws.freeze_panes="B7"
    wb.active.sheet_properties.tabColor="C8962A"
    out=io.BytesIO(); wb.save(out); out.seek(0)
    return out

@app.route('/')
def index(): return render_template('index.html')

@app.route('/parse', methods=['POST'])
def parse():
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
                combined_text+="\n\n=== FILE: {} ===\n".format(fname)+extract_text(fpath)
            except Exception as e:
                return jsonify({"error":"Failed to read {}: {}".format(fname,str(e))}),500
            finally:
                if os.path.exists(fpath): os.remove(fpath)
        else: return jsonify({"error":"Unsupported file: {}".format(file.filename)}),400

    if not combined_text.strip(): return jsonify({"error":"No text extracted"}),400

    try: parsed=parse_with_claude(combined_text,company_name)
    except Exception as e: return jsonify({"error":"AI parsing failed: {}".format(str(e))}),500

    return jsonify(parsed)

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.get_json()
        if not data: return jsonify({"error":"No data received"}),400

        positions = data.get("current_positions",[])
        balances = data.get("balances",{})
        total = 0
        for pos in positions:
            lender = pos.get("lender","")
            bal = float(balances.get(lender, 0) or 0)
            pos["outstanding_balance"] = bal
            total += bal
        data["total_current_positions"] = total

        excel = build_excel(data)
        safe = re.sub(r'[^\w\s-]','',data.get("company_name","analysis")).strip().replace(' ','_')
        return send_file(excel,as_attachment=True,download_name=safe+"_analysis.xlsx",
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error":"Excel generation failed: {}".format(str(e))}),500

if __name__=='__main__':
    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT', 5001)))
