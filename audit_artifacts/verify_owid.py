from pathlib import Path
import requests,re,hashlib,json,openpyxl,pandas as pd
base=Path('audit_artifacts');html=(base/'live_owid_html').read_text()
paths=re.findall(r'"(\/api\/v1\/indicators\/1210053\.[^"]+\.(?:data|metadata)\.json)"',html)
print('paths',paths)
files=[]
for p in dict.fromkeys(paths):
 u='https://api.ourworldindata.org'+p
 r=requests.get(u,timeout=60);fn=p.rsplit('/',1)[-1];(base/f'owid_{fn}').write_bytes(r.content);print(fn,r.status_code,len(r.content),hashlib.sha256(r.content).hexdigest(),r.headers.get('content-type'));files.append((fn,r.content))
for fn,cont in files:
 try:
  obj=json.loads(cont);print('\nJSON',fn, type(obj), (list(obj)[:30] if isinstance(obj,dict) else 'len='+str(len(obj))))
 except Exception as e: print(e)
# parse metadata/data likely entityKey/values
D=json.loads(next(c for fn,c in files if '.data.' in fn)); M=json.loads(next(c for fn,c in files if '.metadata.' in fn))
print('metadata',json.dumps(M,indent=2)[:12000])
print('data sample',json.dumps(D,indent=2)[:2000])
# decode data structure per OWID indicator API typical {values:[],entityKey:[],years:[]?}
# investigate
print('D keys',D.keys())
# Determine known data mapping
