# fanbase_engine.py — SSG Fan Base (In-Unit) bucket engine
# Port of classifyAt() from Visitor-Competitor-Dashboard index.html (~L1442).
# ONE brain: any bucket-rule change happens HERE and in the browser engine together.
#
# Inputs (same folder):
#   visitor.xlsx        FINAL sheet (competitor visits)
#   locations.json      unit universe [{u,b,p,f}] (exported from EMBEDDED_LOCATIONS)
#   fanbase_inputs.json D365 pull from Pull_FanBase_Inputs.ps1
# Output: fanbase.json
#
# Buckets (mutually exclusive, sum = units): A+ / A / B / D / E / F.
# C = OVERLAY (latest won > 90 days) — reported separately, NEVER in the sum.
import json, os, datetime, openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
P = lambda f: os.path.join(HERE, f)
WON_GATE = 'Gate 04.13 Closed As Won'

def canon_service(s):
    u = (s or '').upper()
    if 'POOL' in u or 'SWIM' in u: return 'POOL'
    if 'MAINT' in u or 'HANDYMAN' in u or 'AMC' in u: return 'MAINT'
    if 'CLEAN' in u or 'HOUSE' in u: return 'HOUSEKEEP'
    if 'LAUNDR' in u: return 'LAUNDRY'
    if 'FIT' in u and 'OUT' in u: return 'FITOUT'
    if 'INSPECT' in u or 'SURVEY' in u: return 'INSPECT'
    return 'OTHER'

def is_our_company(name):
    n = (name or '').strip().lower()
    return n.startswith('candoo') or n.startswith('s & c') or n.startswith('s&c') or n.startswith('strive')

def as_list(v):
    # PowerShell ConvertTo-Json collapses 1-element arrays to plain strings
    if not v: return []
    return [v] if isinstance(v, str) else list(v)

def iso_week(d):
    y, w, _ = d.isocalendar()
    return '%d-W%02d' % (y, w)

def main():
    cutoff = datetime.date.today()
    aging = (cutoff - datetime.timedelta(days=90)).isoformat()
    cut = cutoff.isoformat()

    locs = json.load(open(P('locations.json'), encoding='utf-8'))
    inputs = json.load(open(P('fanbase_inputs.json'), encoding='utf-8-sig'))
    proj_inputs = {k.upper(): v for k, v in inputs.get('projects', {}).items()}

    ws = openpyxl.load_workbook(P('visitor.xlsx'), read_only=True)['FINAL']
    it = ws.iter_rows(values_only=True); next(it)
    vms = {}
    vms_projects = set()
    latest_visit = ''
    for r in it:
        date, _typ, purpose, company, _scope, bu, proj = r[0], r[1], r[2], r[3], r[4], r[5], r[6]
        if not bu or not date: continue
        ds = str(date)[:10]
        if ds > cut: continue
        vms_projects.add((proj or '').strip().upper())
        if ds > latest_visit: latest_visit = ds
        if is_our_company(company): continue
        e = vms.setdefault(str(bu).strip(), {'comp': set(), 'n': 0})
        e['comp'].add(canon_service(purpose)); e['n'] += 1

    units_by_project = {}
    for l in locs:
        p = (l.get('p') or '').strip().upper()
        if p: units_by_project.setdefault(p, []).append(l['u'])

    out = {'generatedUtc': datetime.datetime.utcnow().isoformat(timespec='seconds') + 'Z',
           'cutoffDate': cut, 'week': iso_week(cutoff),
           'latestVisitDate': latest_visit,
           'oursSource': inputs.get('source', 'd365'), 'oursPulledAt': inputs.get('pulledAt', ''),
           'projects': {}}

    for proj in sorted(vms_projects):
        units = units_by_project.get(proj, [])
        if not units: continue
        pin = proj_inputs.get(proj, {})
        opps = pin.get('opps', {}) or {}
        books = pin.get('bookings', {}) or {}
        buckets = {'A+': 0, 'A': 0, 'B': 0, 'D': 0, 'E': 0, 'F': 0}
        overlay_c = 0
        for u in units:
            v = vms.get(u, {'comp': set(), 'n': 0})
            bk = books.get(u) or {}
            op = opps.get(u) or {}
            has_bk = (bk.get('count') or 0) > 0
            has_comp = len(v['comp']) > 0
            has_won = (op.get('w') or 0) > 0
            has_any_opp = (op.get('t') or 0) > 0
            has_ours = has_bk or has_any_opp
            if not has_ours and not has_comp: b = 'F'
            elif not has_ours and has_comp:   b = 'E'
            elif has_any_opp and not has_won and not has_bk: b = 'D'
            else:
                sc_serv = set(canon_service(s) for s in as_list(bk.get('services')))
                for d in as_list(op.get('wd')): sc_serv.add(canon_service(d))
                if not has_comp: b = 'A+'
                else: b = 'A' if (sc_serv & v['comp']) else 'B'
            buckets[b] += 1
            if has_won and op.get('lw') and op['lw'] < aging: overlay_c += 1
        fans = buckets['A+'] + buckets['A'] + buckets['B']
        out['projects'][proj] = {
            'units': len(units), 'buckets': buckets, 'overlayC': overlay_c,
            'fans': fans, 'fansPct': round(100.0 * fans / len(units), 1) if units else None}

    json.dump(out, open(P('fanbase.json'), 'w', encoding='utf-8'), indent=1)
    print('WROTE fanbase.json  cutoff=%s week=%s oursSource=%s' % (cut, out['week'], out['oursSource']))
    for p, d in out['projects'].items():
        print('  %-18s units=%-4d %s  C-overlay=%d fans=%d (%.1f%%)' %
              (p, d['units'], d['buckets'], d['overlayC'], d['fans'], d['fansPct']))
    for p, d in out['projects'].items():
        assert sum(d['buckets'].values()) == d['units'], 'SUM MISMATCH ' + p
    print('sanity OK: bucket sums == unit counts')

if __name__ == '__main__':
    main()
