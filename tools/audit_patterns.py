import zipfile,re,sys,json,collections
from lxml import etree
M='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
inv=json.load(open(sys.argv[2]))
files={'Stores':'1_Stores.xlsx','Production':'2_Production.xlsx','Dispatch':'3_Dispatch.xlsx','MIS':'4_MIS.xlsx'}
out={}
for wb,fn in files.items():
    z=zipfile.ZipFile(sys.argv[1]+'/'+fn)
    tabs={t['sheet']:t['name'] for t in inv[wb]['tables'] if t['sheet']}
    W=collections.Counter(); per={}
    for s in inv[wb]['sheets']:
        tn=tabs.get(s['sheet']); c=collections.Counter()
        # need shared formula expansion: count master formula pattern * number of children sharing same si
        si_master={}; si_count=collections.Counter()
        cells=[]
        for ev,el in etree.iterparse(z.open(s['part'])):
            if el.tag==M+'c':
                f=el.find(M+'f')
                if f is not None:
                    if f.get('t')=='shared':
                        si=f.get('si'); si_count[si]+=1
                        if f.text: si_master[si]=f.text
                    else:
                        cells.append(f.text or '')
                el.clear()
        allf=cells+[si_master.get(si,'') for si,n in si_count.items() for _ in range(n)]
        for t in allf:
            fam=set(re.findall(r'(?:_xlfn\.|_xlws\.)*([A-Z][A-Z0-9\.]*)\(',t))
            for x in fam: c['fn:'+x]+=1
            if re.search(r'INDEX\([^()]*\[[^\]]+\],1\):',t): c['running_range']+=1
            if tn and re.search(r'(SUMIFS|COUNTIFS|COUNTIF|SUMIF|MAXIFS)\(\s*'+re.escape(tn)+r'\[',t): c['self_table_agg']+=1
            if re.search(r'ROW\(\)-ROW\(',t): c['row_based_id']+=1
            if 'ANCHORARRAY' in t or '#' in t and re.search(r'\$?[A-Z]+\$?\d+#',t): c['spill_ref']+=1
            if re.search(r'\b(FILTER|SORTBY|UNIQUE|SEQUENCE|SORT)\(',t.replace('_xlfn.','').replace('_xlws.','')): c['dynamic_array']+=1
            if re.search(r'\b(XLOOKUP|XMATCH|LET|LAMBDA|TEXTJOIN|MAXIFS|MINIFS|IFS|SWITCH|CONCAT)\(',t.replace('_xlfn.','').replace('_xlws.','')): c['post2016_fn']+=1
            if re.search(r'\b(TODAY|NOW|OFFSET|INDIRECT|CELL|RAND)\(',t): c['volatile']+=1
            if re.search(r'q[A-Z][A-Za-z]+\[',t): c['reads_import_table']+=1
            c['total']+=1
        per[s['sheet']]=dict(c); W.update(c)
    out[wb]={'total':dict(W),'per':per}
    keys=['total','post2016_fn','dynamic_array','spill_ref','running_range','self_table_agg','row_based_id','volatile','reads_import_table']
    print(wb,{k:W[k] for k in keys})
    for sh,c in per.items():
        if c.get('total',0)>1000: print('   ',sh,{k:c.get(k,0) for k in keys})
json.dump(out,open(sys.argv[3],'w'),indent=1)
