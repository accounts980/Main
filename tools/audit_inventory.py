import zipfile, re, sys, json, collections
from lxml import etree
NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
M='{%s}'%NS['m']
FUNC_RE=re.compile(r'(?:_xlfn\.|_xlws\.)?([A-Z][A-Z0-9\.]*)\(')
def colnum(s):
    n=0
    for ch in s:
        n=n*26+ord(ch)-64
    return n
def run(path,tag):
    z=zipfile.ZipFile(path)
    wb=etree.fromstring(z.read('xl/workbook.xml'))
    rels=etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rmap={r.get('Id'):r.get('Target') for r in rels}
    ss=[]
    if 'xl/sharedStrings.xml' in z.namelist():
        for si in etree.fromstring(z.read('xl/sharedStrings.xml')).findall(M+'si'):
            ss.append(''.join(si.itertext()))
    out=[]
    for sh in wb.find(M+'sheets'):
        name=sh.get('name'); state=sh.get('state','visible'); rid=sh.get('{%s}id'%NS['r'])
        t=rmap[rid]; t=t.lstrip('/'); t=t if t.startswith('xl/') else 'xl/'+t
        info=dict(wb=tag,sheet=name,state=state,part=t,cells=0,values=0,formulas=0,shared_children=0,array_f=0,funcs=collections.Counter(),
                  xlfn=collections.Counter(),fullcol=0,ext_refs=0,maxrow=0,maxcol=0,dim='',dv=0,cf=0,cf_ranges=[],prot=False,freeze=False,
                  rows_with_formula=set(),cols_formula=collections.Counter(),sample_formulas={},hyperlinks=0,mailto=0, cells_text={}, volatile=collections.Counter(), lastvalrow=0)
        data=z.read(t)
        for ev,el in etree.iterparse(__import__('io').BytesIO(data),events=('end',)):
            tg=el.tag
            if tg==M+'c':
                info['cells']+=1
                ref=el.get('r'); m=re.match(r'([A-Z]+)(\d+)',ref); col,row=m.group(1),int(m.group(2))
                f=el.find(M+'f'); v=el.find(M+'v'); isv=el.find(M+'is')
                if f is not None:
                    info['formulas']+=1
                    info['rows_with_formula'].add(row); info['cols_formula'][col]+=1
                    if f.get('t')=='shared' and not f.text: info['shared_children']+=1
                    if f.get('t')=='array': info['array_f']+=1
                    if f.text:
                        txt=f.text
                        for fn in FUNC_RE.findall(txt): info['funcs'][fn]+=1
                        for fn in re.findall(r'_xl(?:fn|ws)\.([A-Z\.]+)',txt): info['xlfn'][fn]+=1
                        if re.search(r"(?<![A-Z\$])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![0-9A-Z])",txt): info['fullcol']+=1
                        if re.search(r"\[\d+\]",txt): info['ext_refs']+=1
                        for vf in ('INDIRECT','OFFSET','NOW','TODAY','RAND','RANDBETWEEN','CELL','INFO'):
                            if re.search(r'\b'+vf+r'\(',txt): info['volatile'][vf]+=1
                        if 'mailto' in txt.lower() or 'HYPERLINK' in txt: info['mailto']+=('mailto' in txt.lower())
                        if col not in info['sample_formulas'] and len(info['sample_formulas'])<60: info['sample_formulas'][col]=(ref,txt[:400])
                if v is not None or isv is not None:
                    info['values']+=1
                    if f is None: info['lastvalrow']=max(info['lastvalrow'],row)
                if row<=12 and (v is not None or isv is not None):
                    val = ss[int(v.text)] if (el.get('t')=='s' and v is not None) else (''.join(isv.itertext()) if isv is not None else (v.text if v is not None else ''))
                    info['cells_text'][ref]=(val or '')[:80]
                info['maxrow']=max(info['maxrow'],row); info['maxcol']=max(info['maxcol'],colnum(col))
                el.clear()
            elif tg==M+'dimension': info['dim']=el.get('ref')
            elif tg==M+'dataValidation': info['dv']+=1
            elif tg==M+'conditionalFormatting':
                info['cf']+=len(el.findall(M+'cfRule')); info['cf_ranges'].append(el.get('sqref'))
            elif tg==M+'sheetProtection': info['prot']=True
            elif tg==M+'pane':
                if el.get('state') in ('frozen','frozenSplit'): info['freeze']=True
            elif tg==M+'hyperlink': info['hyperlinks']+=1
        # x14 dataValidations in extLst
        info['dv_x14']=len(re.findall(rb'<x14:dataValidation ',data))
        info['cf_x14']=len(re.findall(rb'<x14:cfRule ',data))
        info['rows_with_formula']=len(info['rows_with_formula'])
        out.append(info)
    # tables
    tables=[]
    for n in z.namelist():
        if n.startswith('xl/tables/table') and n.endswith('.xml'):
            t=etree.fromstring(z.read(n))
            cols=[c.get('name') for c in t.find(M+'tableColumns')]
            calc=[c.get('name') for c in t.find(M+'tableColumns') if c.find(M+'calculatedColumnFormula') is not None]
            tables.append(dict(part=n,name=t.get('name'),ref=t.get('ref'),cols=cols,calc=calc,tableType=t.get('tableType')))
    # map tables to sheets
    t2s={}
    for info in out:
        rp=info['part'].replace('worksheets/','worksheets/_rels/')+'.rels'
        if rp in z.namelist():
            for r in etree.fromstring(z.read(rp)):
                tgt=r.get('Target')
                if 'tables/' in tgt: t2s['xl/tables/'+tgt.split('tables/')[1]]=info['sheet']
                if 'comments' in tgt: info['comments']=True
                if 'printerSettings' in tgt: info['printer']=True
                if 'drawing' in tgt and 'vml' not in tgt: info['drawing']=True
    for t in tables: t['sheet']=t2s.get(t['part'])
    dn=[(d.get('name'),d.get('localSheetId'),d.text) for d in wb.find(M+'definedNames')] if wb.find(M+'definedNames') is not None else []
    conns=z.read('xl/connections.xml').decode() if 'xl/connections.xml' in z.namelist() else ''
    return dict(sheets=out,tables=tables,names=dn,connections=conns,workbookprot=b'workbookProtection' in z.read('xl/workbook.xml'),parts=z.namelist())
res={}
for p,tag in [('1_Stores.xlsx','Stores'),('2_Production.xlsx','Production'),('3_Dispatch.xlsx','Dispatch'),('4_MIS.xlsx','MIS')]:
    r=run(sys.argv[1]+'/'+p,tag)
    for s in r['sheets']:
        s['funcs']=dict(s['funcs']); s['xlfn']=dict(s['xlfn']); s['cols_formula']=dict(s['cols_formula']); s['volatile']=dict(s['volatile'])
    res[tag]=r
json.dump(res,open(sys.argv[2],'w'),indent=1,default=str)
for tag,r in res.items():
    print('=====',tag,'sheets',len(r['sheets']),'formulas',sum(s['formulas'] for s in r['sheets']),'tables',len(r['tables']),'names',len(r['names']),'wbprot',r['workbookprot'])
    for s in r['sheets']:
        print(f"{s['sheet'][:34]:34} {s['state'][:6]:6} dim={s['dim']:12} maxR={s['maxrow']:6} lastV={s['lastvalrow']:6} cells={s['cells']:7} f={s['formulas']:7} fRows={s['rows_with_formula']:6} dv={s['dv']}+{s['dv_x14']} cf={s['cf']}+{s['cf_x14']} prot={int(s['prot'])} frz={int(s['freeze'])} full={s['fullcol']} vol={s['volatile']} xlfn={dict(sorted(s['xlfn'].items(),key=lambda x:-x[1])[:6])}")
