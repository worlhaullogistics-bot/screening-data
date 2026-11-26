"""Builds small lookup files from two public U.S. government lists, so a Cloudflare Worker on the free plan
can check one company in a few milliseconds:
  - Consolidated Screening List (Commerce, State, Treasury: Entity List, Denied Persons, OFAC SDN, ...)
  - FMC list of active Ocean Transportation Intermediaries (licensed NVOCCs / ocean freight forwarders, registered foreign NVOCCs)
Output: data/meta.json, data/csl/<abc>.json, data/fmc/id/<nn>.json, data/fmc/name/<abc>.json
The matching rules (norm, tokens) must stay identical to the worker's (worldhaul-setup worker/src.js)."""
import csv, io, json, os, re, sys, unicodedata, urllib.request, datetime, collections, shutil
csv.field_size_limit(10**9)
CSL_URL = 'https://data.trade.gov/downloadable_consolidated_screening_list/v1/consolidated.csv'
FMC_URL = 'https://www2.fmc.gov/OTI/reports/OTI_List.xlsx'
OUT = sys.argv[1] if len(sys.argv) > 1 else 'data'
STOP = set('LLC L L C INC INCORPORATED LTD LIMITED CO COMPANY CORP CORPORATION SA SAS SRL SPA AG GMBH BV NV PLC PTE PVT LP LLP DE CV RL S A THE AND OF Y E DBA'.split())
GENERIC = set('LOGISTICS SHIPPING TRADING TRADE INTERNATIONAL INTL FREIGHT CARGO EXPRESS SERVICES SERVICE TRANSPORT TRANSPORTATION LINE LINES GLOBAL AIR SEA OCEAN GROUP HOLDING HOLDINGS ENTERPRISES ENTERPRISE INDUSTRIES INDUSTRY IMPORT EXPORT FORWARDING FORWARDERS FORWARDER AGENCY AGENCIES WORLDWIDE CORPORATE MARINE MARITIME SUPPLY CHAIN SOLUTIONS NETWORK TECHNOLOGY TECHNOLOGIES TECH INSTITUTE STOCK JOINT ELECTRONICS ELECTRONIC SCIENCE SCIENTIFIC RESEARCH DEVELOPMENT ENGINEERING EQUIPMENT MACHINERY INDUSTRIAL MANUFACTURING DEFENSE DEFENCE INVESTMENT INVESTMENTS BANK OIL GAS ENERGY CONSTRUCTION PETROLEUM CHEMICAL CHEMICALS MATERIALS METAL METALS SYSTEMS PLANT FACTORY STATE NATIONAL UNIVERSITY'.split())
def norm(s):
    s = unicodedata.normalize('NFD', str(s or '')); s = ''.join(c for c in s if not unicodedata.combining(c))
    s = s.upper().replace('&', ' AND ')
    return re.sub(r'\s+', ' ', re.sub(r'[^A-Z0-9 ]', ' ', s)).strip()
def tokens(s): return [t for t in norm(s).split() if t not in STOP]
def keys(s):  # shard keys for a name: first 3 letters of each distinctive word
    ks = {t[:3].lower() for t in tokens(s) if len(t) >= 3 and t not in GENERIC and not t.isdigit()}
    if not ks:
        ks = {t[:3].lower() for t in tokens(s) if len(t) >= 3 and not t.isdigit()}
    return ks
def get(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'screening-data builder (github.com/worlhaullogistics-bot/screening-data)'})
    with urllib.request.urlopen(req, timeout=180) as r: return r.read()
def short_src(s):
    s = str(s or '')
    for a, b in [('Specially Designated Nationals', 'OFAC SDN'), ('Entity List', 'BIS Entity List'), ('Denied Persons', 'BIS Denied Persons'), ('Unverified List', 'BIS Unverified'), ('Military End User', 'BIS Military End User'), ('Nonproliferation Sanctions', 'State Nonproliferation'), ('ITAR Debarred', 'State ITAR Debarred'), ('Sectoral Sanctions', 'OFAC Sectoral'), ('Foreign Sanctions Evaders', 'OFAC Foreign Sanctions Evaders'), ('Palestinian Legislative', 'OFAC NS-PLC'), ('Menu-Based', 'OFAC NS-MBS'), ('Chinese Military', 'OFAC NS-CMIC'), ('Capta', 'OFAC CAPTA')]:
        if a.lower() in s.lower(): return b
    return s[:40]
def build_csl(raw):
    text = raw.decode('utf-8-sig', 'replace'); rd = csv.DictReader(io.StringIO(text))
    cols = {c.lower().strip(): c for c in rd.fieldnames}
    col = lambda *names: next((cols[n] for n in names if n in cols), None)
    cN, cS, cP, cA, cAlt, cId, cT, cC = col('name'), col('source'), col('programs'), col('addresses'), col('alt_names'), col('_id', 'id', 'entity_number'), col('type'), col('citizenships', 'nationalities')
    print('CSL columns:', rd.fieldnames[:30])
    shards, n = collections.defaultdict(dict), 0
    for row in rd:
        name = (row.get(cN) or '').strip()
        if not name: continue
        n += 1
        addr = row.get(cA) or ''
        countries = sorted({p.strip().split(',')[-1].strip() for p in addr.split(';') if p.strip()} - {''})[:6]
        ent = {'n': name, 's': short_src(row.get(cS)), 'p': (row.get(cP) or '')[:120], 'c': countries, 't': (row.get(cT) or '')[:20], 'id': (row.get(cId) or '')[:40]}
        alts = [a.strip() for a in (row.get(cAlt) or '').split(';') if a.strip()][:25]
        for nm in [name] + alts:
            e = dict(ent); 
            if nm != name: e['a'] = nm
            for k in keys(nm): shards[k][(e['id'], nm)] = e
    return shards, n
def build_fmc(raw):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    ids, names, n = collections.defaultdict(dict), collections.defaultdict(dict), 0
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True); header = None
        for r in rows:
            cells = [str(c or '').strip() for c in r]
            if header is None:
                low = [c.lower() for c in cells]
                if any('organization' in c for c in low) and any('name' in c for c in low): header = low; print('FMC sheet', ws.title, 'columns:', cells); continue
                continue
            d = dict(zip(header, cells))
            pick = lambda *subs: next((v for k, v in d.items() if v and all(s in k for s in subs)), '')
            org = re.sub(r'\D', '', pick('organization', 'no') or pick('org', 'number'))
            name = pick('organization', 'name') or pick('name')
            if not (org or name): continue
            lic = re.sub(r'[^0-9A-Z]', '', (pick('license') or '').upper())
            typ = ' '.join(v for k, v in d.items() if v and ('type' in k or 'oti' in k or 'business' in k))[:80]
            e = {'org': org, 'lic': lic, 'n': name, 'tn': pick('trade')[:200], 'city': pick('city'), 'st': pick('state'), 'cty': pick('country'),
                 'type': typ or ws.title, 'bond': (pick('bond') or pick('surety'))[:80], 'sheet': ws.title}
            n += 1
            for idv in {org.lstrip('0'), lic.lstrip('0')} - {''}:
                ids[idv[-2:].zfill(2)].setdefault(idv, [])
                ids[idv[-2:].zfill(2)][idv].append(e)
            for nm in [name] + [t.strip() for t in re.split(r'[;|]', e['tn']) if t.strip()]:
                for k in keys(nm): names[k][(org, nm)] = dict(e, m=nm)
    return ids, names, n
def write(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w') as f: json.dump(obj, f, separators=(',', ':'), ensure_ascii=False, sort_keys=True)
def main():
    tmp = OUT + '.new'; shutil.rmtree(tmp, ignore_errors=True); meta = {'built': datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'), 'sources': {'csl': CSL_URL, 'fmc': FMC_URL}}
    csl_raw = open(os.environ['CSL_FILE'], 'rb').read() if os.environ.get('CSL_FILE') else get(CSL_URL)
    shards, n = build_csl(csl_raw); meta['csl_entries'] = n; meta['csl_shards'] = len(shards)
    for k, v in shards.items(): write(f'{tmp}/csl/{k}.json', sorted(v.values(), key=lambda e: (e.get('a') or e['n'], e['id'])))
    if n < 1000: sys.exit(f'Too few sanctions entries ({n}); keeping the previous data.')
    # FMC: the site's own download when it allows it; otherwise the copy a person downloaded from www2.fmc.gov/oti
    # ("Download OTI List") and uploaded to sources/OTI_List.xlsx in this repository.
    m = 0
    try:
        if os.environ.get('FMC_FILE'): fmc_raw = open(os.environ['FMC_FILE'], 'rb').read()
        else:
            try: fmc_raw = get(FMC_URL)
            except Exception as e:
                print('FMC download not allowed from here:', e)
                fmc_raw = open('sources/OTI_List.xlsx', 'rb').read(); print('Using sources/OTI_List.xlsx (downloaded by hand)')
                meta['fmc_source'] = 'sources/OTI_List.xlsx, uploaded ' + datetime.datetime.fromtimestamp(os.path.getmtime('sources/OTI_List.xlsx'), datetime.timezone.utc).strftime('%Y-%m-%d')
        ids, names, m = build_fmc(fmc_raw)
    except Exception as e:
        print('FMC list not available:', e)
    if m >= 1000:
        for k, v in ids.items(): write(f'{tmp}/fmc/id/{k}.json', v)
        for k, v in names.items(): write(f'{tmp}/fmc/name/{k}.json', sorted(v.values(), key=lambda e: (e['m'], e['org'])))
        meta['fmc_entries'] = m; meta['fmc_built'] = meta['built']
    elif os.path.isdir(f'{OUT}/fmc'):  # keep the last good FMC files
        shutil.copytree(f'{OUT}/fmc', f'{tmp}/fmc')
        old = json.load(open(f'{OUT}/meta.json')) if os.path.exists(f'{OUT}/meta.json') else {}
        meta['fmc_entries'] = old.get('fmc_entries', 0); meta['fmc_built'] = old.get('fmc_built')
    else: meta['fmc_entries'] = 0
    write(f'{tmp}/meta.json', meta)
    shutil.rmtree(OUT, ignore_errors=True); os.rename(tmp, OUT)
    sizes = sorted((os.path.getsize(os.path.join(dp, f)), os.path.join(dp, f)) for dp, _, fs in os.walk(OUT) for f in fs)
    print(json.dumps(meta), 'files', len(sizes), 'largest', sizes[-3:])
if __name__ == '__main__': main()
