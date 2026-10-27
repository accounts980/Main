import zipfile,re,sys,json
from lxml import etree
M='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
inv=json.load(open(sys.argv[2]))
files={'Stores':'1_Stores.xlsx','Production':'2_Production.xlsx','Dispatch':'3_Dispatch.xlsx','MIS':'4_MIS.xlsx'}
def cn(s):
    n=0
    for ch in s: n=n*26+ord(ch)-64
    return n
def cl(n):
    s=''
    while n: n,r=divmod(n-1,26); s=chr(65+r)+s
    return s
res=[]
for wb,fn in files.items():
    z=zipfile.ZipFile(sys.argv[1]+'/'+fn)
    part={s['sheet']:s['part'] for s in inv[wb]['sheets']}
    for t in inv[wb]['tables']:
        if t['sheet'] is None: continue
        a,b=t['ref'].split(':'); ma=re.match(r'([A-Z]+)(\d+)',a); mb=re.match(r'([A-Z]+)(\d+)',b)
        c0=cn(ma.group(1)); hr=int(ma.group(2)); last=int(mb.group(2))
        inputs=[c for c in t['cols'] if c not in t['calc']]
        key=None
        for cand in ['Date','Count Date','Entry Date','JWC No','Challan No','Material','Supplier','Item Name','Product','Name','Machine','Customer','Vendor','Coater','Model','Grade','Category','Reason','Where Used','Operator','Key','Product / Job','Rejected At','BOM Model','Lot (RM ID)']:
            if cand in t['cols']: key=cand;break
        if key is None: key=t['cols'][0]
        col=cl(c0+t['cols'].index(key))
        cnt=0
        for ev,el in etree.iterparse(z.open(part[t['sheet']])):
            if el.tag==M+'c':
                r=el.get('r'); m=re.match(r'([A-Z]+)(\d+)',r)
                if m.group(1)==col and hr<int(m.group(2))<=last:
                    v=el.find(M+'v'); isv=el.find(M+'is')
                    if (v is not None and v.text not in (None,'')) or isv is not None: cnt+=1
                el.clear()
        body=last-hr
        res.append(dict(wb=wb,sheet=t['sheet'],table=t['name'],ref=t['ref'],cols=len(t['cols']),calc=len(t['calc']),body_rows=body,data_rows=cnt,key=key))
        print(f"{wb:10} {t['sheet'][:26]:26} {t['name']:18} {t['ref']:12} cols={len(t['cols']):3} calc={len(t['calc']):3} body={body:6} data={cnt:6} key={key}")
json.dump(res,open(sys.argv[3],'w'),indent=1)
