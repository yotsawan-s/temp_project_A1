from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule

# ---------- palette ----------
NAVY="1F3864"; BLUE="2E5496"; LTBLUE="D6E0F0"; GREY="F2F2F2"
YELLOW="FFF2CC"; INPUT="FFFF00"; GREEN="C6EFCE"; GREENT="006100"
RED="FFC7CE"; REDT="9C0006"; ORANGE="FFEB9C"; ORANGET="9C5700"
WHITE="FFFFFF"
F="Arial"
thin=Side(style="thin",color="BFBFBF")
border=Border(left=thin,right=thin,top=thin,bottom=thin)

def st(c,*,b=False,sz=10,color="000000",fill=None,wrap=False,h="left",v="center",bd=False):
    c.font=Font(name=F,bold=b,size=sz,color=color)
    if fill: c.fill=PatternFill("solid",fgColor=fill)
    c.alignment=Alignment(horizontal=h,vertical=v,wrap_text=wrap)
    if bd: c.border=border

def title(ws,text,sub,ncol):
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncol)
    st(ws.cell(1,1,text),b=True,sz=15,color=WHITE,fill=NAVY,h="left"); ws.row_dimensions[1].height=26
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=ncol)
    st(ws.cell(2,1,sub),b=False,sz=10,color=WHITE,fill=BLUE,h="left",wrap=True); ws.row_dimensions[2].height=30

def hdr(ws,row,headers,start=1,fill=BLUE,color=WHITE):
    for i,h in enumerate(headers):
        st(ws.cell(row,start+i,h),b=True,sz=10,color=color,fill=fill,h="center",wrap=True,bd=True)
    ws.row_dimensions[row].height=28

wb=Workbook()

# ============================================================ README
ws=wb.active; ws.title="README"
title(ws,"PART RECONCILE  —  BOT Reconciliation Tool  (Prototype v0.1)",
      "วัตถุประสงค์: ตรวจสอบว่าข้อมูล Transaction ใน A1_System (ไฟล์ 1_A / 1_B) ตรงกับ Product Database (D / O / C) ครบถ้วนถูกต้อง ก่อนส่งต่อหน่วยงานอื่น  |  ตรวจตาม Data Transaction Date ที่ผู้ใช้ระบุ",6)
ws.column_dimensions["A"].width=3
ws.column_dimensions["B"].width=34
for c in "CDEF": ws.column_dimensions[c].width=22
r=4
def line(text,**kw):
    global r
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=6)
    st(ws.cell(r,2,text),**kw); r+=1

line("1)  แนวคิดการ Reconcile (2 มุมมอง)",b=True,sz=12,color=NAVY)
line("ทุก Set จะจับคู่ระหว่าง  Product Database  ⇄  A1_System  โดยใช้ Key Mapping เป็นตัวเชื่อม แล้วตรวจ 2 ทิศทาง:")
line("    • มุมที่ 1  A1_System → Database : A1_System บันทึกเกินหรือไม่ (มีใน A1 แต่ไม่มีใน Database = EXTRA)")
line("    • มุมที่ 2  Database → A1_System : A1_System บันทึกขาดหรือไม่ (มีใน Database แต่ไม่มีใน A1 = MISSING)")
line("    • เมื่อ Key ตรงกันทั้งคู่ จะตรวจค่า (จำนวนเงิน/Amount) ต่อ ว่าตรงกันไหม (ไม่ตรง = AMOUNT MISMATCH)")
line("ผลลัพธ์ที่เป็น Error จะถูกสรุปเป็น \"ERROR Message\" ผูกกับ Key เพื่อนำกลับไปแก้ไขที่ระบบ A1_System")
line("(สอดคล้องกับไฟล์ 1_Bxxx ใน mockup ที่มีคอลัมน์ = [Key , ERROR Message] ซึ่งคือไฟล์ผลลัพธ์ของการ Reconcile)",sz=9,color="808080")
r+=1
line("2)  โครงสร้างไฟล์ (จากชีท 0_Structure / 1_Flow_Map / 3_DataSource)",b=True,sz=12,color=NAVY)
line("    • A1_System  :  1_A001–1_A009 (9 ไฟล์)  +  1_B001–1_B009 (9 ไฟล์ผลลัพธ์/ERROR)")
line("    • Database   :  D = 2_D001–2_D005 , O = 3_O001–3_O003 , C = 4_C001–4_C003")
line("    • Reconcile รวม 21 Sets : Product D = 12 , Product O = 6 , Product C = 3")
r+=1
line("3)  วิธีใช้งาน (Prototype)",b=True,sz=12,color=NAVY)
line("    1. ไปที่ชีท 'Control_Panel' กรอกช่องสีเหลือง: Data Transaction Date, Department Code, Product ที่จะรัน")
line("    2. (Prototype) ชีท 'Sample_Data' จำลองข้อมูลของ Set ตัวอย่าง D1-1 (2_D001 ⇄ 1_A009)")
line("    3. ชีท 'Recon_Detail' แสดงผลการตรวจรายบรรทัด + ERROR Message (คำนวณสดด้วยสูตร)")
line("    4. ชีท 'Recon_Summary' = Dashboard สรุปทุก Set ว่า OK / มี Error กี่รายการ")
r+=1
line("4)  คำอธิบายสี (Legend)",b=True,sz=12,color=NAVY)
legend=[("ช่องกรอกข้อมูล (Input)",INPUT,"000000"),
        ("OK / ตรงกัน (Matched)",GREEN,GREENT),
        ("MISSING in A1 — บันทึกขาด",RED,REDT),
        ("EXTRA in A1 — บันทึกเกิน",RED,REDT),
        ("AMOUNT MISMATCH — ค่าไม่ตรง",ORANGE,ORANGET)]
for txt,fl,co in legend:
    st(ws.cell(r,2,"  "),fill=fl,bd=True); st(ws.cell(r,3,txt),color=co,b=True); r+=1
r+=1
line("5)  ข้อสมมุติ & คำถามที่รอผู้ใช้ยืนยัน  (ดูชีท Open_Questions)",b=True,sz=12,color="9C0006")
ws.sheet_view.showGridLines=False

# ============================================================ CONTROL PANEL
ws=wb.create_sheet("Control_Panel")
title(ws,"CONTROL PANEL  —  พารามิเตอร์การรัน",
      "กรอกช่องสีเหลือง แล้วระบบจะตรวจ Transaction ตามวันที่/หน่วยงานที่ระบุ  (ค่าเหล่านี้ใช้เป็นตัวกรอง Col.Filter ของทุกไฟล์)",4)
ws.sheet_view.showGridLines=False
for c,w in zip("ABCD",[3,30,26,40]): ws.column_dimensions[c].width=w
inputs=[("Data Transaction Date","2026-05-30","วันที่ Transaction ที่ต้องการตรวจ (ตรงกับ Data Set Date / TRADE_DATE / Settlement Date ฯลฯ)"),
        ("Department Code","0123","รหัสหน่วยงาน (ตัวกรอง Dept Code / DEPT CODE)"),
        ("Product to run","ALL","D / O / C / ALL"),
        ("Amount tolerance","0","ผลต่างจำนวนเงินที่ยอมรับได้ (0 = ต้องตรงเป๊ะ)")]
r=4
st(ws.cell(r,2,"พารามิเตอร์"),b=True,color=WHITE,fill=BLUE,bd=True)
st(ws.cell(r,3,"ค่า (กรอก)"),b=True,color=WHITE,fill=BLUE,bd=True,h="center")
st(ws.cell(r,4,"คำอธิบาย"),b=True,color=WHITE,fill=BLUE,bd=True); r+=1
names={}
for lab,val,desc in inputs:
    st(ws.cell(r,2,lab),b=True,bd=True)
    cell=ws.cell(r,3,val); st(cell,fill=INPUT,bd=True,h="center")
    st(ws.cell(r,4,desc),wrap=True,bd=True,sz=9)
    names[lab]=f"Control_Panel!$C${r}"; r+=1
# named ranges
from openpyxl.workbook.defined_name import DefinedName
wb.defined_names.add(DefinedName("ChkDate",attr_text=names["Data Transaction Date"]))
wb.defined_names.add(DefinedName("Tol",attr_text=names["Amount tolerance"]))
r+=1
st(ws.cell(r,2,"สถานะผลรวม (Overall Status)"),b=True,sz=12,color=NAVY); r+=1
st(ws.cell(r,2,"จำนวน Error ของ Set ตัวอย่าง (D1-1)"),b=True,bd=True)
ws.cell(r,3,'=COUNTIF(Recon_Detail!$G:$G,"OK / Matched")')  # placeholder, fixed below
errcell=f"$C${r}"
st(ws.cell(r,3),bd=True,h="center")
r0=r
# We'll set real formula after detail known; set now referencing Recon_Detail status col G
ws.cell(r,3).value='=COUNTA(Recon_Detail!$H$6:$H$100)-COUNTIF(Recon_Detail!$H$6:$H$100,"")'

# ============================================================ MAP CONFIG
ws=wb.create_sheet("Map_Config")
title(ws,"MAP CONFIG  —  แผนผังการ Reconcile (จากชีท 1_Flow_Map)",
      "แต่ละ Set = สายการตรวจ Source#1 → Source#4  พร้อม Key Mapping & คอลัมน์วันที่ของแต่ละไฟล์ (ดึงจาก 3_DataSource)  |  Trans.Map = cardinality เช่น 1:1:n",13)
ws.sheet_view.showGridLines=False
keymap={'3_O001':'FI Arrangement Number','1_A007':'FI Arrangement Number','1_A005':'FI Arrangement Number',
'1_A001':'FI Arrangement Number','1_A002':'Arrangement Number','3_O003':'CompositeOptSeqNum','3_O002':'OptionNum',
'1_A008':'Reference Transaction Number','4_C001':'Transaction No + DR/CR','4_C002':'Ccy + Notional','4_C003':'Currency + Accrued Int',
'2_D001':'TRADE_REF','1_A009':'FI Arrangement Number + CMF CODE','2_D002':'FI Arrangement Number','1_A003':'FI Arrangement Number',
'1_A004':'Arr Number','1_A006':'FI Arrangement Number','2_D004':'(none)','2_D005':'(none)',
'1_B001':'Arrangement Number','1_B002':'Fi Arrangement Number','1_B003':'Fi Arrangement Number','1_B004':'Arrangement Number',
'1_B005':'Fi Arrangement Number','1_B006':'Fi Arrangement Number','1_B007':'Fi Arrangement Number','1_B008':'Ref.No.',
'1_B009':'Fi Arrangement Number','2_D003':'Ref + CCY + Amount + CMF Code'}
datemap={'3_O001':'Arrangement Contract Date','1_A007':'Data Set Date','1_A005':'Data Set Date','1_A001':'Data Set Date',
'1_A002':'Data Set Date','3_O003':'Date','3_O002':'(none)','1_A008':'Data Set Date','4_C001':'(none)','4_C002':'Settlement Date',
'4_C003':'Settlement Date','2_D001':'TRADE_DATE','1_A009':'Data Set Date','2_D002':'Data Set Date','1_A003':'Data Set Date',
'1_A004':'Data Set Date','1_A006':'Data Set Date','2_D004':'(none)','2_D005':'(none)','2_D003':'Date'}
for b in ['1_B001','1_B002','1_B003','1_B004','1_B005','1_B006','1_B007','1_B008','1_B009']:
    datemap[b]='(none)'
sets=[("D","D1",1,["2_D001","1_A009","1_B009"],"1 : 1 : n"),
("D","D1",2,["2_D001","1_A005","1_B002"],"1 : 1 : n"),
("D","D1",3,["2_D001","1_A001","1_B003"],"1 : 1,2 : n"),
("D","D1",4,["2_D001","1_A002","1_B004"],"1 : 1 : n"),
("D","D2",5,["2_D002","1_A003","2_D001","1_B006"],"1 : 1 : 1 : n"),
("D","D3",6,["2_D003","1_A001","1_B003"],"1 : 1 : n"),
("D","D3",7,["2_D003","1_A002","1_B004"],"1 : 1 : n"),
("D","D3",8,["2_D003","1_A004","1_B001"],"1 : 1 : n"),
("D","D3",9,["2_D003","1_A005","1_B002"],"1 : 1 : n"),
("D","D3",10,["2_D003","1_A008","1_B008"],"1 : 1 : n"),
("D","D3",11,["2_D003","1_A006","1_B005"],"1 : n : n"),
("D","D3",12,["2_D003","2_D004","2_D005","1_A005"],"n : 1 : 1 : ?"),
("O","O1",1,["3_O001","1_A007","1_B007"],"1 : 1 : n"),
("O","O1",2,["3_O001","1_A005","1_B002"],"1 : 1 : n"),
("O","O1",3,["3_O001","1_A001","3_O003","1_B003"],"1 : 1 : 1 : n"),
("O","O1",4,["3_O001","1_A002","1_B004"],"1 : 1 : n"),
("O","O2",5,["3_O002","1_A005","1_B002"],"1 : 2 : n"),
("O","O3",6,["3_O003","1_A008","1_B008"],"1 : 1 : n"),
("C","C1",1,["1_A008","4_C001"],"1 : 2"),
("C","C1",2,["1_A008","4_C002","4_C003"],"1 : 1 : 1"),
("C","C1",3,["1_A008","1_B008"],"1 : n")]
heads=["Product","Category","Set No.","Source#1","Source#2","Source#3","Source#4","Trans.Map",
       "Key Mapping (Source#1)","Date Filter (Source#1)","Database file","A1_System file(s)","Output (1_B / ERROR)"]
hdr(ws,4,heads)
ws.freeze_panes="A5"
r=5
for p,cat,no,chain,tm in sets:
    s=chain+[""]*(4-len(chain))
    db=[x for x in chain if x[0] in "234"]
    a1=[x for x in chain if x.startswith("1_A")]
    out=[x for x in chain if x.startswith("1_B")]
    row=[p,cat,no,s[0],s[1],s[2],s[3],tm,keymap.get(chain[0],""),datemap.get(chain[0],""),
         ", ".join(db),", ".join(a1) or "-",", ".join(out) or "(no 1_B — value match only)"]
    for i,v in enumerate(row):
        c=ws.cell(r,i+1,v); st(c,bd=True,sz=9,h="center" if i in(0,1,2,7) else "left")
    st(ws.cell(r,1),b=True,fill=GREY,h="center")
    r+=1
widths=[8,9,7,11,11,11,11,13,24,18,16,16,26]
for i,w in enumerate(widths): ws.column_dimensions[get_column_letter(i+1)].width=w

# ============================================================ SOURCE CONFIG
ws=wb.create_sheet("Source_Config")
title(ws,"SOURCE CONFIG  —  สรุปเฉพาะ Field ที่จำเป็นของแต่ละไฟล์ (จากชีท 3_DataSource)",
      "แสดงเฉพาะ Flag = Y :  Date Filter (กรองตามวันที่)  +  Key Mapping (คีย์เชื่อม)  +  จำนวนคอลัมน์ที่ Select ไว้แสดงผล",6)
ws.sheet_view.showGridLines=False
selcount={'3_O001':15,'1_A007':6,'1_A005':9,'1_A001':9,'1_A002':6,'3_O003':6,'3_O002':13,'1_A008':6,'4_C001':0,
'4_C002':2,'4_C003':2,'2_D001':15,'1_A009':12,'2_D002':52,'1_A003':52,'1_A004':6,'1_A006':5,'2_D004':5,'2_D005':5,
'1_B001':2,'1_B002':2,'1_B003':2,'1_B004':2,'1_B005':2,'1_B006':2,'1_B007':2,'1_B008':2,'1_B009':2,'2_D003':8}
group={'1_A':'A1_System (main)','1_B':'A1_System (ERROR output)','2_D':'Database D','3_O':'Database O','4_C':'Database C'}
def grp(s):
    for k,v in group.items():
        if s.startswith(k): return v
    return ""
hdr(ws,4,["Group","Source / File","Date Filter Column","Key Mapping Column(s)","# Columns Selected","Role in Reconcile"])
ws.freeze_panes="A5"
order=['2_D001','2_D002','2_D003','2_D004','2_D005','3_O001','3_O002','3_O003','4_C001','4_C002','4_C003',
'1_A001','1_A002','1_A003','1_A004','1_A005','1_A006','1_A007','1_A008','1_A009',
'1_B001','1_B002','1_B003','1_B004','1_B005','1_B006','1_B007','1_B008','1_B009']
role={'1_A':'ระบบที่ตรวจ (ตรวจเกิน/ขาด)','1_B':'ไฟล์ผลลัพธ์ ERROR (key+message)','2_D':'ฐานข้อมูลต้นทาง (D)','3_O':'ฐานข้อมูลต้นทาง (O)','4_C':'ฐานข้อมูลต้นทาง (C)'}
r=5
for s in order:
    row=[grp(s),s,datemap.get(s,""),keymap.get(s,""),selcount.get(s,""),next((v for k,v in role.items() if s.startswith(k)),"")]
    for i,v in enumerate(row):
        st(ws.cell(r,i+1,v),bd=True,sz=9,h="center" if i in(4,) else "left")
    if s.startswith("1_B"): st(ws.cell(r,2),b=True,fill=YELLOW)
    r+=1
for i,w in enumerate([20,16,24,30,16,30]): ws.column_dimensions[get_column_letter(i+1)].width=w

# ============================================================ SAMPLE DATA
ws=wb.create_sheet("Sample_Data")
title(ws,"SAMPLE DATA  —  ข้อมูลจำลองของ Set ตัวอย่าง D1-1   (2_D001  ⇄  1_A009)",
      "Prototype: ข้อมูลสมมุติเพื่อสาธิตการคำนวณสด  |  สมมุติว่า TRADE_REF (Database) = FI Arrangement Number (A1_System)  |  กรองแล้วตาม Date/Dept",8)
ws.sheet_view.showGridLines=False
# DB table 2_D001
st(ws.cell(4,2,"ตาราง A : Product Database  2_D001  (กรองตาม TRADE_DATE = ChkDate)"),b=True,color=WHITE,fill="538135")
ws.merge_cells("B4:E4")
hdr(ws,5,["TRADE_REF (Key)","CCY","RCV_INIT_PRINCIPAL (Amount)","TRADE_DATE"],start=2,fill="A9D08E",color="000000")
db_rows=[("FX001","USD",1000000,"2026-05-30"),("FX002","USD",2000000,"2026-05-30"),
("FX003","EUR",1500000,"2026-05-30"),("FX004","JPY",3000000,"2026-05-30")]
r=6
for k,ccy,amt,dt in db_rows:
    st(ws.cell(r,2,k),bd=True); st(ws.cell(r,3,ccy),bd=True,h="center")
    c=ws.cell(r,4,amt); st(c,bd=True,h="right"); c.number_format="#,##0"
    st(ws.cell(r,5,dt),bd=True,h="center"); r+=1
db_last=r-1
# A1 table 1_A009
st(ws.cell(4,7,"ตาราง B : A1_System  1_A009  (กรองตาม Data Set Date = ChkDate, Dept)"),b=True,color=WHITE,fill="2E75B6")
ws.merge_cells("G4:J4")
hdr(ws,5,["FI Arrangement Number (Key)","Currency","Initial Buy Amount","Data Set Date"],start=7,fill="9DC3E6",color="000000")
a1_rows=[("FX001","USD",1000000,"2026-05-30"),("FX002","USD",2050000,"2026-05-30"),
("FX003","EUR",1500000,"2026-05-30"),("FX005","GBP",800000,"2026-05-30")]
r=6
for k,ccy,amt,dt in a1_rows:
    st(ws.cell(r,7,k),bd=True); st(ws.cell(r,8,ccy),bd=True,h="center")
    c=ws.cell(r,9,amt); st(c,bd=True,h="right"); c.number_format="#,##0"
    st(ws.cell(r,10,dt),bd=True,h="center"); r+=1
a1_last=r-1
note=("หมายเหตุข้อมูลจำลอง:  FX004 มีใน Database แต่ไม่มีใน A1 (บันทึกขาด/MISSING)  |  "
"FX005 มีใน A1 แต่ไม่มีใน Database (บันทึกเกิน/EXTRA)  |  FX002 มีทั้งคู่แต่จำนวนเงินไม่ตรง (MISMATCH)")
ws.merge_cells("B12:J13"); st(ws.cell(12,2,note),wrap=True,color="9C5700",fill=YELLOW,bd=True)
for col,w in zip("BCDEFGHIJ",[18,8,24,14,3,28,12,18,14]): ws.column_dimensions[col].width=w
DBK=f"Sample_Data!$B$6:$B${db_last}"; DBA=f"Sample_Data!$D$6:$D${db_last}"
A1K=f"Sample_Data!$G$6:$G${a1_last}"; A1A=f"Sample_Data!$I$6:$I${a1_last}"

# ============================================================ RECON DETAIL
ws=wb.create_sheet("Recon_Detail")
title(ws,"RECON DETAIL  —  ผลการตรวจรายบรรทัด  (Set D1-1 : 2_D001 ⇄ 1_A009)",
      "ตรวจ 2 ทิศทางพร้อมกันต่อ Key  |  สูตรคำนวณสดจากชีท Sample_Data  |  คอลัมน์ ERROR Message = ข้อความนำกลับไปแก้ที่ A1_System (ผูกกับไฟล์ 1_B009)",8)
ws.sheet_view.showGridLines=False
hdr(ws,5,["Key (Arrangement No.)","In Database? (2_D001)","In A1_System? (1_A009)",
          "DB Amount","A1 Amount","Amount Diff","Result / Status","ERROR Message (→ fix in A1_System)"])
ws.freeze_panes="A6"
# union of keys (prototype: hardcode union list, presence/amount/status by formula)
union=["FX001","FX002","FX003","FX004","FX005"]
r=6
for k in union:
    ws.cell(r,1,k); st(ws.cell(r,1),bd=True,b=True)
    ws.cell(r,2,f'=IF(COUNTIF({DBK},A{r})>0,"YES","NO")'); st(ws.cell(r,2),bd=True,h="center")
    ws.cell(r,3,f'=IF(COUNTIF({A1K},A{r})>0,"YES","NO")'); st(ws.cell(r,3),bd=True,h="center")
    ws.cell(r,4,f'=SUMIF({DBK},A{r},{DBA})'); st(ws.cell(r,4),bd=True,h="right"); ws.cell(r,4).number_format="#,##0"
    ws.cell(r,5,f'=SUMIF({A1K},A{r},{A1A})'); st(ws.cell(r,5),bd=True,h="right"); ws.cell(r,5).number_format="#,##0"
    ws.cell(r,6,f'=D{r}-E{r}'); st(ws.cell(r,6),bd=True,h="right"); ws.cell(r,6).number_format="#,##0;(#,##0)"
    ws.cell(r,7,(f'=IF(AND(B{r}="YES",C{r}="NO"),"EXTRA in A1",'
                 f'IF(AND(B{r}="NO",C{r}="YES"),"MISSING in A1",'
                 f'IF(ABS(F{r})>Tol,"AMOUNT MISMATCH","OK / Matched")))'))
    st(ws.cell(r,7),bd=True,b=True,h="center")
    ws.cell(r,8,(f'=IF(G{r}="EXTRA in A1","พบใน A1_System แต่ไม่มีใน Database — ตรวจสอบ/ลบรายการเกินที่ A1_System",'
                 f'IF(G{r}="MISSING in A1","มีใน Database แต่ขาดใน A1_System — บันทึกเพิ่มที่ A1_System",'
                 f'IF(G{r}="AMOUNT MISMATCH","จำนวนเงินไม่ตรง (DB="&TEXT(D{r},"#,##0")&" / A1="&TEXT(E{r},"#,##0")&") — แก้จำนวนเงินที่ A1_System","")))'))
    st(ws.cell(r,8),bd=True,wrap=True,color=REDT)
    r+=1
last=r-1
for col,w in zip("ABCDEFGH",[20,18,20,14,14,13,18,48]): ws.column_dimensions[col].width=w
# conditional formatting on status col G
rng=f"A6:H{last}"
ws.conditional_formatting.add(rng,FormulaRule(formula=[f'$G6="OK / Matched"'],fill=PatternFill("solid",fgColor=GREEN),font=Font(name=F,color=GREENT)))
ws.conditional_formatting.add(rng,FormulaRule(formula=[f'OR($G6="EXTRA in A1",$G6="MISSING in A1")'],fill=PatternFill("solid",fgColor=RED),font=Font(name=F,color=REDT,bold=True)))
ws.conditional_formatting.add(rng,FormulaRule(formula=[f'$G6="AMOUNT MISMATCH"'],fill=PatternFill("solid",fgColor=ORANGE),font=Font(name=F,color=ORANGET,bold=True)))

# ============================================================ RECON SUMMARY
ws=wb.create_sheet("Recon_Summary",2)  # place after Control_Panel
title(ws,"RECON SUMMARY  —  Dashboard สรุปทุก Set",
      "หนึ่งแถวต่อหนึ่ง Set  |  สีแดง = พบ Error ต้องแก้  |  Set D1-1 คำนวณสดจากข้อมูลจริง, Set อื่นรอโหลดข้อมูล (Pending)",9)
ws.sheet_view.showGridLines=False
hdr(ws,4,["Product","Category","Set No.","Chain (Database ⇄ A1_System)","Trans.Map",
          "MISSING in A1","EXTRA in A1","AMOUNT MISMATCH","Overall Status"])
ws.freeze_panes="A5"
r=5
for p,cat,no,chain,tm in sets:
    chain_s=" ⇄ ".join(chain)
    for i,v in enumerate([p,cat,no,chain_s,tm]):
        st(ws.cell(r,i+1,v),bd=True,sz=9,h="center" if i in(0,1,2,4) else "left")
    if (p,cat,no)==("D","D1",1):
        ws.cell(r,6,f'=COUNTIF(Recon_Detail!$G$6:$G${last},"MISSING in A1")')
        ws.cell(r,7,f'=COUNTIF(Recon_Detail!$G$6:$G${last},"EXTRA in A1")')
        ws.cell(r,8,f'=COUNTIF(Recon_Detail!$G$6:$G${last},"AMOUNT MISMATCH")')
        ws.cell(r,9,f'=IF(SUM(F{r}:H{r})=0,"OK","ERROR — "&SUM(F{r}:H{r})&" รายการ")')
        for cc in range(6,10): st(ws.cell(r,cc),bd=True,h="center",b=(cc==9))
    else:
        for cc in range(6,9):
            st(ws.cell(r,cc,"-"),bd=True,h="center")
        st(ws.cell(r,9,"Pending data"),bd=True,h="center",color="808080")
    r+=1
slast=r-1
for col,w in zip("ABCDEFGHI",[8,9,7,40,13,13,12,15,18]): ws.column_dimensions[col].width=w
rng=f"I5:I{slast}"
ws.conditional_formatting.add(rng,FormulaRule(formula=['$I5="OK"'],fill=PatternFill("solid",fgColor=GREEN),font=Font(name=F,color=GREENT,bold=True)))
ws.conditional_formatting.add(rng,FormulaRule(formula=['LEFT($I5,5)="ERROR"'],fill=PatternFill("solid",fgColor=RED),font=Font(name=F,color=REDT,bold=True)))

# ============================================================ OPEN QUESTIONS
ws=wb.create_sheet("Open_Questions")
title(ws,"OPEN QUESTIONS  —  จุดที่ต้องการข้อมูลเพิ่มเพื่อทำเวอร์ชันจริง",
      "Prototype ตั้งสมมุติฐานไว้ตามด้านล่าง โปรดยืนยัน/แก้ไข เพื่อพัฒนาเป็น Production tool",4)
ws.sheet_view.showGridLines=False
for c,w in zip("ABCD",[3,5,52,40]): ws.column_dimensions[c].width=w
hdr(ws,4,["","#","คำถาม / ประเด็น","สมมุติฐานปัจจุบันใน Prototype"],start=1)
qs=[("1","Key Mapping ข้ามไฟล์: DB ใช้คีย์ของตัวเอง (เช่น TRADE_REF, OptionNum) แต่ A1 ใช้ FI Arrangement Number — มีตารางแปลงคีย์ (translation) หรือถือว่าเป็นค่าเดียวกัน?","สมมุติว่าเป็นคีย์ธุรกิจเดียวกัน (TRADE_REF = FI Arrangement No.)"),
("2","Cardinality (Trans.Map เช่น 1:1:n, 1:n:n): ฝั่ง 'n' ต้องรวมยอด (SUM) แล้วเทียบ หรือเทียบรายบรรทัด?","สมมุติ: รวมยอด (SUMIF) ตามคีย์แล้วเทียบ"),
("3","ตรวจแค่การมีอยู่ของรายการ (existence) หรือเทียบค่าด้วย? ถ้าเทียบค่า ใช้ field ใด (Amount/Notional/...)?","เทียบทั้ง existence + จำนวนเงิน 1 field"),
("4","วันที่ตรวจ: ใช้วันที่เดียว (เป๊ะ) หรือช่วงวันที่? และแต่ละไฟล์ใช้คอลัมน์วันที่ต่างกัน ถูกต้องไหม?","วันที่เดียว = ChkDate, map คอลัมน์ตาม Source_Config"),
("5","ไฟล์ผลลัพธ์ 1_Bxxx: ให้เครื่องมือ 'เขียน' ERROR Message กลับไปไฟล์เหล่านี้เลยไหม หรือแค่แสดงในรายงาน?","Prototype แสดงในรายงาน (คอลัมน์ ERROR Message)"),
("6","Chain ที่ยาว 3-4 ขั้น (เช่น D2-5, O1-3): ตรวจแบบ end-to-end (ต้นทาง↔ปลายทาง) หรือทีละคู่ต่อขั้น?","ยังไม่ทำใน prototype — รอแนวทาง"),
("7","Product C (1_A008 ⇄ 4_C0xx) ใช้คีย์ผสม (Ccy+Notional / Transaction No+DR/CR) — ยืนยัน logic การ match","รอตัวอย่างข้อมูลจริง"),
("8","มีข้อมูลตัวอย่างจริง (ไฟล์ใดไฟล์หนึ่ง) ให้ลองได้ไหม เพื่อ map คอลัมน์ให้ตรง","ใช้ข้อมูลจำลองใน Sample_Data")]
r=5
for n,q,a in qs:
    st(ws.cell(r,2,n),b=True,bd=True,h="center")
    st(ws.cell(r,3,q),wrap=True,bd=True,sz=9)
    st(ws.cell(r,4,a),wrap=True,bd=True,sz=9,fill=YELLOW,color="9C5700")
    ws.row_dimensions[r].height=46; r+=1

# order sheets
order_sheets=["README","Control_Panel","Recon_Summary","Recon_Detail","Sample_Data","Map_Config","Source_Config","Open_Questions"]
wb._sheets.sort(key=lambda s: order_sheets.index(s.title))
wb.calculation.fullCalcOnLoad=True
wb.save("Part_Reconcile_Tool_Prototype.xlsx")
print("saved")
