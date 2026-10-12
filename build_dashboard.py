import pandas as pd, json, os
H=os.path.dirname(os.path.abspath(__file__))
d=pd.read_csv(f"{H}/data/clean_sales.csv",parse_dates=["OrderDate"])
d["ym"]=d.OrderDate.dt.strftime("%Y-%m")
r=lambda x:round(float(x),0)
M=[[a,b,c,r(rv),r(pf),int(n)] for (a,b,c),(rv,pf,n) in d.groupby(["ym","Region","Category"]).agg({"Revenue":"sum","Profit":"sum","OrderID":"count"}).iterrows()] if False else \
  [[k[0],k[1],k[2],r(v.Revenue),r(v.Profit),int(v.OrderID)] for k,v in d.groupby(["ym","Region","Category"]).agg({"Revenue":"sum","Profit":"sum","OrderID":"count"}).iterrows()]
P=[[int(k[0]),k[1],k[2],k[3],r(v.Revenue),r(v.Profit)] for k,v in d.groupby(["Year","Region","Category","Product"])[["Revenue","Profit"]].sum().iterrows()]
D=[[int(k[0]),k[1],k[2],float(k[3]),r(v.Revenue),r(v.Profit)] for k,v in d.groupby(["Year","Region","Category","Discount"])[["Revenue","Profit"]].sum().iterrows()]
html=open(f"{H}/dashboard_template.html").read().replace("__DATA__",json.dumps({"M":M,"P":P,"D":D},separators=(",",":")))
open(f"{H}/Sales_Dashboard_PowerBI_Style.html","w").write(html)
print(len(html)//1024,"KB")
