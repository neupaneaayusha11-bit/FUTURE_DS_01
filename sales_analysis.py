"""Business Sales Performance Analytics - Future Interns Task 1
Steps: 1) create sample dataset  2) clean  3) analyse  4) build Excel dashboard
NOTE: Dataset is SIMULATED (task allows self-created data). Replace data/raw_sales.csv
with Superstore / Online Retail data and adjust column names to use real data."""
import numpy as np, pandas as pd, os
from openpyxl import Workbook
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUT = os.path.dirname(os.path.abspath(__file__))
rng = np.random.default_rng(42)

# ---------- 1. SIMULATED RAW DATA ----------
PRODUCTS = {  # product: (category, list price, cost ratio)
 "Laptop":("Technology",900,.78),"Smartphone":("Technology",600,.75),"Monitor":("Technology",250,.72),
 "Headphones":("Technology",120,.60),"Keyboard":("Technology",45,.55),
 "Office Chair":("Furniture",220,.82),"Desk":("Furniture",450,.85),"Bookcase":("Furniture",180,.80),"Sofa":("Furniture",700,.88),
 "Paper":("Office Supplies",25,.50),"Binder":("Office Supplies",8,.40),"Pens":("Office Supplies",12,.45),
 "Printer Ink":("Office Supplies",60,.60),"Stapler":("Office Supplies",15,.50)}
names = list(PRODUCTS)
pw = np.array([.05,.05,.04,.05,.06, .05,.03,.04,.02, .15,.12,.12,.07,.06]); pw/=pw.sum()
days = pd.date_range("2023-01-01","2025-12-31")
season = {1:.7,2:.75,3:.9,4:.95,5:1,6:1,7:.95,8:1,9:1.1,10:1.15,11:1.5,12:1.6}
dw = np.array([season[d.month]*(1.18**(d.year-2023)) for d in days]); dw/=dw.sum()
N = 6000
df = pd.DataFrame({"OrderID":[f"ORD-{i:05d}" for i in range(1,N+1)],
  "OrderDate":rng.choice(days,N,p=dw),
  "Region":rng.choice(["West","East","South","North"],N,p=[.32,.28,.22,.18]),
  "Segment":rng.choice(["Consumer","Corporate","Home Office"],N,p=[.5,.3,.2]),
  "Product":rng.choice(names,N,p=pw)})
df["Category"]=df.Product.map(lambda p:PRODUCTS[p][0])
def qty(c): return rng.integers(1,4) if c=="Technology" else rng.integers(1,5) if c=="Furniture" else rng.integers(2,16)
df["Quantity"]=df.Category.map(qty)
levels=[0,.05,.1,.2,.3]
def disc(r):
    if r.Category=="Furniture": p=[.25,.15,.2,.2,.2]
    else: p=[.45,.2,.15,.12,.08]
    if r.Region=="South": p=[p[0]*.6,p[1],p[2],p[3]*1.4,p[4]*1.8]
    p=np.array(p)/sum(p); return rng.choice(levels,p=p)
df["Discount"]=df.apply(disc,axis=1)
df["UnitPrice"]=[round(PRODUCTS[p][1]*rng.uniform(.95,1.05),2) for p in df.Product]
df["UnitCost"]=[round(PRODUCTS[p][1]*PRODUCTS[p][2],2) for p in df.Product]
# make it "dirty"
raw=df.copy()
raw["OrderDate"]=[d.strftime("%Y-%m-%d") if i%3 else d.strftime("%d/%m/%Y") for i,d in enumerate(raw.OrderDate)]
raw["Region"]=raw.Region.astype(object)
for i in rng.choice(N,120,replace=False): raw.at[i,"Region"]=raw.at[i,"Region"].upper()
for i in rng.choice(N,120,replace=False): raw.at[i,"Region"]=" "+raw.at[i,"Region"].lower()+" "
for i in rng.choice(N,90,replace=False): raw.at[i,"Region"]=np.nan
for i in rng.choice(N,40,replace=False): raw.at[i,"Quantity"]=-int(raw.at[i,"Quantity"])
raw=pd.concat([raw,raw.sample(180,random_state=1)]).sample(frac=1,random_state=2).reset_index(drop=True)
raw.to_csv(f"{OUT}/data/raw_sales.csv",index=False)

# ---------- 2. CLEANING ----------
log=[("Raw rows",len(raw))]
d=raw.copy()
n=len(d); d=d.drop_duplicates(); log.append(("Duplicate rows removed",n-len(d)))
d["Region"]=d.Region.str.strip().str.title()
n=len(d); d=d.dropna(subset=["Region"]); log.append(("Rows with missing Region removed",n-len(d)))
n=len(d); d=d[d.Quantity>0]; log.append(("Rows with invalid (<=0) Quantity removed",n-len(d)))
dt=pd.to_datetime(d.OrderDate,format="%Y-%m-%d",errors="coerce")
dt=dt.fillna(pd.to_datetime(d.OrderDate,format="%d/%m/%Y",errors="coerce"))
d["OrderDate"]=dt; d=d.dropna(subset=["OrderDate"])
d["Month"]=d.OrderDate.dt.to_period("M").dt.to_timestamp(); d["Year"]=d.OrderDate.dt.year
d["Revenue"]=(d.Quantity*d.UnitPrice*(1-d.Discount)).round(2)
d["Cost"]=(d.Quantity*d.UnitCost).round(2); d["Profit"]=(d.Revenue-d.Cost).round(2)
d=d.sort_values("OrderDate").reset_index(drop=True); log.append(("Clean rows",len(d)))
d.to_csv(f"{OUT}/data/clean_sales.csv",index=False)

# ---------- 3. ANALYSIS (for written insights) ----------
R=d.Revenue.sum(); P=d.Profit.sum()
cat=d.groupby("Category")[["Revenue","Profit"]].sum(); cat["Margin"]=cat.Profit/cat.Revenue
reg=d.groupby("Region")[["Revenue","Profit"]].sum(); reg["Margin"]=reg.Profit/reg.Revenue
prod=d.groupby("Product")[["Revenue","Profit"]].sum().sort_values("Revenue",ascending=False); prod["Margin"]=prod.Profit/prod.Revenue
yr=d.groupby("Year").Revenue.sum()
mon=d.groupby(d.OrderDate.dt.month).Revenue.sum()
q4=mon[[10,11,12]].sum()/mon.sum()
d["DiscBand"]=d.Discount
db=d.groupby("DiscBand")[["Revenue","Profit"]].sum(); db["Margin"]=db.Profit/db.Revenue
fur30=d[(d.Category=="Furniture")&(d.Discount>=.2)]
pc=lambda x:f"{x:.1%}"; mn=lambda x:f"{x:,.0f}"
ins=[
 f"Total revenue is {mn(R)} with profit {mn(P)} (overall margin {pc(P/R)}) from {len(d):,} order lines.",
 f"Revenue grew {pc(yr[2024]/yr[2023]-1)} in 2024 and {pc(yr[2025]/yr[2024]-1)} in 2025 - steady growth.",
 f"{cat.Revenue.idxmax()} is the top category by revenue ({pc(cat.Revenue.max()/R)} of sales); {cat.Margin.idxmax()} has the best margin ({pc(cat.Margin.max())}).",
 f"{cat.Margin.idxmin()} has the weakest margin ({pc(cat.Margin.min())}) - discounting is hurting it.",
 f"{reg.Revenue.idxmax()} is the top region ({pc(reg.Revenue.max()/R)} of revenue); {reg.Margin.idxmin()} has the lowest margin ({pc(reg.Margin.min())}) because of heavier discounts.",
 f"Seasonality: Oct-Dec (Q4) brings {pc(q4)} of yearly revenue; best month is {pd.Timestamp(2000,mon.idxmax(),1):%B}, weakest is {pd.Timestamp(2000,mon.idxmin(),1):%B}.",
 f"Top 3 products ({', '.join(prod.index[:3])}) give {pc(prod.Revenue[:3].sum()/R)} of revenue.",
 f"Orders with 20%+ discount earn only {pc(db.Margin[db.index>=.2].sum()/1 if False else d[d.Discount>=.2].Profit.sum()/d[d.Discount>=.2].Revenue.sum())} margin vs {pc(d[d.Discount==0].Profit.sum()/d[d.Discount==0].Revenue.sum())} with no discount.",
 f"Furniture sold at 20%+ discount is {'loss-making' if fur30.Profit.sum()<0 else 'low margin'} (profit {mn(fur30.Profit.sum())}).",
]
rec=[
 "Cap discounts at 10% for Furniture and review pricing on Sofa/Desk - protect margin.",
 "Stock up and run campaigns from September to December; use the slow Jan-Feb period for clearance and bundles.",
 f"Push high-margin Technology and Office Supplies add-ons (e.g. headphones, ink) alongside big-ticket items.",
 f"Fix discount discipline in the {reg.Margin.idxmin()} region with approval rules for discounts above 10%.",
 f"Double down on {reg.Revenue.idxmax()} (largest market) while testing campaigns in the smallest region to grow it.",
 "Track KPIs monthly: revenue, margin %, AOV, discount rate, and top-5 products.",
]
pd.DataFrame({"Insight":ins}).to_csv(f"{OUT}/data/insights.csv",index=False)

# ---------- 4. EXCEL WORKBOOK ----------
F="Arial"; HF=PatternFill("solid",fgColor="1F3864"); 
def hdr(ws,row,cols):
    for c in range(1,cols+1):
        x=ws.cell(row,c); x.font=Font(name=F,bold=True,color="FFFFFF"); x.fill=HF; x.alignment=Alignment(horizontal="center")
wb=Workbook(); dash=wb.active; dash.title="Dashboard"
cd=wb.create_sheet("Clean_Data")
cols=["OrderID","OrderDate","Month","Year","Region","Category","Product","Segment","Quantity","UnitPrice","Discount","Revenue","Cost","Profit"]
cd.append(cols); hdr(cd,1,len(cols))
for i,r in enumerate(d.itertuples(),start=2):
    cd.append([r.OrderID,r.OrderDate.to_pydatetime(),r.Month.to_pydatetime(),r.Year,r.Region,r.Category,r.Product,r.Segment,
               r.Quantity,r.UnitPrice,r.Discount,f"=I{i}*J{i}*(1-K{i})",round(r.Quantity*r.UnitCost,2),f"=L{i}-M{i}"])
last=len(d)+1
for row in cd.iter_rows(min_row=2):
    row[1].number_format="yyyy-mm-dd"; row[2].number_format="mmm-yy"; row[10].number_format="0%"
    for k in (9,11,12,13): row[k].number_format="#,##0.00"
    for x in row: x.font=Font(name=F,size=10)
for i,w in enumerate([11,12,9,7,9,15,13,12,9,10,9,12,12,12],1): cd.column_dimensions[get_column_letter(i)].width=w
cd.freeze_panes="A2"
rg=lambda col:f"Clean_Data!${col}$2:${col}${last}"

def table(name,keys,keycol,fmt_key=None):
    ws=wb.create_sheet(name); ws.append([name.replace("_"," "),"Revenue","Profit","Margin %","Revenue Share","Orders"]); hdr(ws,1,6)
    for i,k in enumerate(keys,start=2):
        ws.cell(i,1,k)
        ws.cell(i,2,f"=SUMIFS({rg('L')},{rg(keycol)},A{i})"); ws.cell(i,3,f"=SUMIFS({rg('N')},{rg(keycol)},A{i})")
        ws.cell(i,4,f"=IFERROR(C{i}/B{i},0)"); ws.cell(i,5,f"=B{i}/SUM($B$2:$B${len(keys)+1})")
        ws.cell(i,6,f"=COUNTIFS({rg(keycol)},A{i})")
        for c in (2,3): ws.cell(i,c).number_format="#,##0"
        for c in (4,5): ws.cell(i,c).number_format="0.0%"
        if fmt_key: ws.cell(i,1).number_format=fmt_key
    for row in ws.iter_rows(min_row=2):
        for x in row: x.font=Font(name=F,size=10)
    ws.column_dimensions["A"].width=20
    for c in "BCDEF": ws.column_dimensions[c].width=14
    return ws
months=sorted(d.Month.unique())
wm=table("Monthly",[pd.Timestamp(m).to_pydatetime() for m in months],"C","mmm-yy")
wm["G1"]="MoM Growth"; hdr(wm,1,7)
for i in range(3,len(months)+2): wm.cell(i,7,f"=IFERROR(B{i}/B{i-1}-1,0)").number_format="0.0%"; wm.cell(i,7).font=Font(name=F,size=10)
wc=table("Category",list(cat.sort_values("Revenue",ascending=False).index),"F")
wr=table("Region",list(reg.sort_values("Revenue",ascending=False).index),"E")
wp=table("Products",list(prod.index),"G")
wy=table("Yearly",[2023,2024,2025],"D"); wy["G1"]="YoY Growth"; hdr(wy,1,7)
for i in (3,4): wy.cell(i,7,f"=B{i}/B{i-1}-1").number_format="0.0%"; wy.cell(i,7).font=Font(name=F,size=10)
# discount table
wd=table("Discount",levels,"K","0%")
# Region x Category margin matrix
wx=wb.create_sheet("Region_x_Category"); cats=list(cat.index); regs=list(reg.index)
wx.append(["Profit margin %"]+cats); hdr(wx,1,len(cats)+1)
for i,rn in enumerate(regs,start=2):
    wx.cell(i,1,rn)
    for j,cn in enumerate(cats,start=2):
        L=get_column_letter(j)
        wx.cell(i,j,f"=IFERROR(SUMIFS({rg('N')},{rg('E')},$A{i},{rg('F')},{L}$1)/SUMIFS({rg('L')},{rg('E')},$A{i},{rg('F')},{L}$1),0)").number_format="0.0%"
for row in wx.iter_rows():
    for x in row: x.font=Font(name=F,size=10,bold=x.row==1 or x.column==1)
wx.column_dimensions["A"].width=16
for c in "BCD": wx.column_dimensions[c].width=17

# Insights sheet
wi=wb.create_sheet("Insights"); wi["A1"]="Key Insights"; wi["A1"].font=Font(name=F,bold=True,size=14)
for i,t in enumerate(ins,start=3): wi.cell(i,1,f"{i-2}. {t}").font=Font(name=F,size=11)
r0=len(ins)+5; wi.cell(r0,1,"Recommendations").font=Font(name=F,bold=True,size=14)
for i,t in enumerate(rec,start=r0+2): wi.cell(i,1,f"{i-r0-1}. {t}").font=Font(name=F,size=11)
wi.cell(r0+len(rec)+3,1,"Note: figures in the text are from the Python analysis of the same Clean_Data sheet; dataset is simulated.").font=Font(name=F,italic=True,size=9)
wi.column_dimensions["A"].width=140

# Cleaning log
wl=wb.create_sheet("Cleaning_Log"); wl.append(["Cleaning step","Rows"]); hdr(wl,1,2)
for a,b in log: wl.append([a,b])
wl.column_dimensions["A"].width=45
for row in wl.iter_rows(min_row=2):
    for x in row: x.font=Font(name=F,size=10)

# Dashboard
dash.sheet_view.showGridLines=False
dash["A1"]="Business Sales Performance Dashboard (2023-2025)"; dash["A1"].font=Font(name=F,bold=True,size=18,color="1F3864")
dash["A2"]="Source: simulated sales data (Clean_Data sheet). All KPIs are live formulas."; dash["A2"].font=Font(name=F,italic=True,size=9)
kp=[("Total Revenue",f"=SUM({rg('L')})","#,##0"),("Total Profit",f"=SUM({rg('N')})","#,##0"),
    ("Profit Margin",f"=B5/A5","0.0%"),("Orders",f"=COUNTA({rg('A')})","#,##0"),
    ("Avg Order Value",f"=A5/D5","#,##0"),("2025 vs 2024",f"=Yearly!G4","0.0%")]
for i,(lab,f,nf) in enumerate(kp):
    c=i*2+1 if False else i+1
    dash.cell(4,c,lab).font=Font(name=F,bold=True,color="FFFFFF",size=10); dash.cell(4,c).fill=HF; dash.cell(4,c).alignment=Alignment(horizontal="center")
    dash.cell(5,c,f).number_format=nf; dash.cell(5,c).font=Font(name=F,bold=True,size=14); dash.cell(5,c).alignment=Alignment(horizontal="center")
dash["C5"]="=B5/A5"; dash["B5"]=f"=SUM({rg('N')})"; dash["A5"]=f"=SUM({rg('L')})"
dash["E5"]="=A5/D5"
for c in "ABCDEF": dash.column_dimensions[c].width=19
def addbar(ws,title,anchor,cats_ref,data_cols,ws_src,n,horizontal=False,w=15,h=7.5):
    ch=BarChart(); ch.type="bar" if horizontal else "col"; ch.title=title; ch.width=w; ch.height=h
    for col in data_cols: ch.add_data(Reference(ws_src,min_col=col,min_row=1,max_row=n+1),titles_from_data=True)
    ch.set_categories(cats_ref); ws.add_chart(ch,anchor); return ch
lc=LineChart(); lc.title="Monthly Revenue Trend"; lc.width=31; lc.height=7.5
lc.add_data(Reference(wm,min_col=2,min_row=1,max_row=len(months)+1),titles_from_data=True)
lc.set_categories(Reference(wm,min_col=1,min_row=2,max_row=len(months)+1)); lc.x_axis.number_format="mmm-yy"
dash.add_chart(lc,"A8")
addbar(dash,"Revenue & Profit by Category","A24",Reference(wc,min_col=1,min_row=2,max_row=4),[2,3],wc,3)
addbar(dash,"Revenue & Profit by Region","D24",Reference(wr,min_col=1,min_row=2,max_row=5),[2,3],wr,4)
addbar(dash,"Revenue by Product","A40",Reference(wp,min_col=1,min_row=2,max_row=15),[2],wp,14,True,15,9)
addbar(dash,"Profit Margin by Discount Level","D40",Reference(wd,min_col=1,min_row=2,max_row=6),[4],wd,5,False,15,9)
for ws in wb.worksheets:
    if ws.title in ("Dashboard","Clean_Data","Insights","Cleaning_Log","Region_x_Category"): continue
    pass
wb.save(f"{OUT}/Sales_Performance_Dashboard.xlsx")
print("\n".join(f"{a}: {b}" for a,b in log)); print("\n".join(ins)); print(); print(db.Margin.round(3).to_dict()); print(cat.round(3)); print(reg.round(3))
