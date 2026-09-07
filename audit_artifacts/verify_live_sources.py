from pathlib import Path
import requests,openpyxl,pandas as pd,hashlib,json,sys,datetime
base=Path('audit_artifacts')
urls={
'MMR':'https://api.worldbank.org/v2/country/all/indicator/SH.STA.MMRT?format=json&per_page=20000',
'GDP':'https://api.worldbank.org/v2/country/all/indicator/NY.GDP.PCAP.PP.KD?format=json&per_page=20000',
'CPR':'https://api.worldbank.org/v2/country/all/indicator/SP.DYN.CONU.ZS?format=json&per_page=20000',
'OWID_HTML':'https://archive.ourworldindata.org/20260805-143540/grapher/political-regime.html',
'OWID_DATA':'https://archive.ourworldindata.org/20260805-143540/grapher/political-regime.json',
'OWID_CSV':'https://archive.ourworldindata.org/20260805-143540/grapher/political-regime.csv',
}
for name,u in urls.items():
 try:
  r=requests.get(u,timeout=45,headers={'User-Agent':'academic-audit/1.0'});print(name,r.status_code,r.url,r.headers.get('content-type'),len(r.content),hashlib.sha256(r.content).hexdigest());(base/f'live_{name.lower()}').write_bytes(r.content)
 except Exception as e:print(name,'ERROR',repr(e))
# WDI verification
p='② 附件二_数据表_Jolie FemTech.xlsx'; wb=openpyxl.load_workbook(p,data_only=True);ws=wb['跨国底稿']
records=[]
for r in range(5,172):
 if ws.cell(r,1).value:
  records.append({'row':r,'iso':ws.cell(r,1).value,'country':ws.cell(r,2).value,'MMR':ws.cell(r,3).value,'MMRyear':ws.cell(r,4).value,'GDP':ws.cell(r,5).value,'GDPyear':ws.cell(r,6).value,'CPR':ws.cell(r,11).value,'CPRyear':ws.cell(r,12).value})
df=pd.DataFrame(records)
for indicator,col,yearcol in [('MMR','MMR','MMRyear'),('GDP','GDP','GDPyear'),('CPR','CPR','CPRyear')]:
 data=json.loads((base/f'live_{indicator.lower()}').read_text())
 meta,obs=data[0],data[1]
 print('\nAPI',indicator,'meta',meta)
 by={}
 for x in obs:
  iso=x['countryiso3code']; y=x['date'];val=x['value']
  if iso and val is not None and y and 2010<=int(y)<=2023:
   by.setdefault(iso,[]).append((int(y),float(val),x['country']['value']))
 ver=[]
 for _,d in df.iterrows():
  cand=by.get(d.iso,[])
  latest=max(cand) if cand else None
  exval=d[col]; exyear=d[yearcol]
  stat='MISSING_IN_LIVE'
  if latest:
   ly,lv,lc=latest
   if exval is None or pd.isna(exval): stat='XLSX_MISSING_LIVE_VALUE'
   elif int(exyear)==ly and abs(float(exval)-lv)<=max(1e-7,abs(lv)*1e-10):stat='MATCH'
   elif int(exyear)==ly: stat='VALUE_DIFFER'
   else:stat='YEAR_DIFFER'
   ver.append({**d.to_dict(),'indicator':indicator,'live_year':ly,'live_value':lv,'live_country':lc,'status':stat,'diff':None if exval is None or pd.isna(exval) else float(exval)-lv})
  else: ver.append({**d.to_dict(),'indicator':indicator,'live_year':None,'live_value':None,'live_country':None,'status':stat,'diff':None})
 out=pd.DataFrame(ver);out.to_csv(base/f'live_wdi_{indicator.lower()}_compare.csv',index=False)
 print('SUMMARY',indicator,out.status.value_counts().to_dict())
 print(out.loc[out.status!='MATCH',['iso','country',col,yearcol,'live_value','live_year','status','diff']].to_string(index=False))
