import openpyxl, datetime, sys, os, re, glob
KEEP = {"LAUNDRY","INSPECTION","CLEANING/CLEANERS","MAINTENANCE/HANDYMAN","FIT-OUT","AMCCONTRACTORS"}
norm = lambda s: str(s or '').upper().replace(' ','')
def resolve(rel):
    h=glob.glob('/sessions/*/mnt/'+rel); return h[0] if h else '/sessions/current/mnt/'+rel
SOURCES=[('BALQIS RESIDENCE',"Balqis Security's files - DAILY CONTRACTORS RECORDS/Visitors Details Balqis Residence.xlsx",None,False),
 ('THE8',"Abdul  Muqeet's files - VMS-DATA FILES/VMS-TH8.xlsx",["VMS-TH8-'26"],True),
 ('NORTH RESIDENCE',"VMS DATA SOUTH & NORTH/VISITOR DATA NORTH RESIDENCE - 2026.xlsx",['NORTH RESIDENCE'],False),
 ('SOUTH RESIDENCE',"VMS DATA SOUTH & NORTH/VISITOR DATA SOUTH RESIDENCE - 2026.xlsx",['SOUTH RESIDENCE'],False)]
OUT=resolve("CRM Related/Visitor-Competitor-Dashboard/visitor.xlsx")
HEADER=['Check In Date','Check In Type','Check In Purpose','Company Name','Scope of work','Building/ Unit','Project Name']
def col(ci,*names):
    low={k.lower():v for k,v in ci.items()}
    for n in names:
        if n.lower() in low: return low[n.lower()]
    return None
def parse_date(v):
    if isinstance(v,(datetime.datetime,datetime.date)):
        dt=v if isinstance(v,datetime.datetime) else datetime.datetime(v.year,v.month,v.day)
    elif isinstance(v,str) and v.strip():
        m=re.match(r'^(\d{2})(\d{2})(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?',v.strip())
        if not m: return None
        dd,mm,yy=int(m.group(1)),int(m.group(2)),int(m.group(3))
        hh=int(m.group(4) or 0);mi=int(m.group(5) or 0);ss=int(m.group(6) or 0)
        try: dt=datetime.datetime(yy,mm,dd,hh,mi,ss)
        except: return None
    else: return None
    if dt.year<2024 or dt.year>2027: return None
    return dt
def clean_source(proj,rel,sheets,dedupe):
    path=resolve(rel)
    if not os.path.exists(path): print('  !! MISSING',proj,file=sys.stderr); return []
    wb=openpyxl.load_workbook(path,read_only=True,data_only=True)
    sns=sheets or wb.sheetnames; out,seen,dropped=[],set(),0
    for sn in sns:
        if sn not in wb.sheetnames: continue
        ws=wb[sn]; it=ws.iter_rows(values_only=True); hdr=None
        for row in it:
            v=[str(c) if c is not None else '' for c in row]
            if 'Check In Date' in v: hdr=v; break
        if not hdr: continue
        ci={h:i for i,h in enumerate(hdr)}
        iD=col(ci,'Check In Date');iT=col(ci,'Check In Type');iP=col(ci,'Check In Purpose')
        iC=col(ci,'Company Name','COMPANY NAME');iS=col(ci,'Scope of work')
        iBU=col(ci,'Building/ Unit','Building/Unit');iPN=col(ci,'Project Name','Project name');iU=col(ci,'Unit')
        for row in it:
            if not row or (iD is not None and row[iD] in (None,'')): continue
            if norm(row[iP] if iP is not None else '') not in KEEP: continue
            comp=str(row[iC]).strip() if iC is not None and row[iC] is not None else ''
            is_dima='dima' in comp.lower()
            typ=str(row[iT]).strip().lower() if iT is not None and row[iT] is not None else ''
            if not is_dima and typ!='unit visit': continue
            date=parse_date(row[iD])
            if date is None: dropped+=1; continue
            pur=str(row[iP]).strip(); comp=comp.upper()
            scope=str(row[iS]).strip() if iS is not None and row[iS] is not None else ''
            bu=str(row[iBU]).strip() if iBU is not None and row[iBU] is not None else ''
            pn=str(row[iPN]).strip() if iPN is not None and row[iPN] is not None else proj
            unit=str(row[iU]).strip() if iU is not None and row[iU] is not None else ''
            if dedupe:
                key=(date.strftime('%Y-%m-%d'),norm(pur),unit.upper(),comp)
                if key in seen: continue
                seen.add(key)
            out.append([date,(typ.title() if typ else 'Unit Visit'),pur,comp,scope,bu,pn])
    print('  %s: %d rows (dropped %d)'%(proj,len(out),dropped)); return out
rows=[]
for p,rel,s,d in SOURCES: rows+=clean_source(p,rel,s,d)
wb=openpyxl.Workbook(); ws=wb.active; ws.title='FINAL'; ws.append(HEADER)
for r in rows: ws.append(r)
wb.save(OUT); print('WROTE %d rows'%len(rows))
