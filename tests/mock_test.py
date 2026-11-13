import json, re, time, datetime, sys
from urllib.parse import urlparse, parse_qs, unquote
from playwright.sync_api import sync_playwright

d = json.load(open('/home/claude/cantina/payload.json'))
LOCS = ['Cantinetta Robin','Cantinetta Nonno','Cantinetta Nonno Cucina','Corridoio Nonno','Bixio','Cantinetta Tadino','Casa Nonno Sardegna','Cantina Corridoni','Cantina Cadamosto']
locs = [dict(id=i+1, nome=n, tipo='fisica', ruolo=None, capacita=30 if n=='Cantinetta Robin' else (120 if n=='Cantinetta Nonno' else None), attiva=True, ordine=i+1, note=None) for i,n in enumerate(LOCS)]
locs.append(dict(id=10, nome='Da ritrovare', tipo='sconosciuta', ruolo='Posizione sconosciuta', capacita=None, attiva=True, ordine=99, note=None))
LID = {l['nome']: l['id'] for l in locs}
cols = ['produttore','vino','annata','formato','denominazione','regione','paese','tipologia','stile','vitigno','dw_inizio','dw_fine','fonte_dw','bere_presto']
wines = [dict(id=i+1, note=None, **dict(zip(cols, w))) for i,w in enumerate(d['w'])]
movs=[]; segn=[]; mid=[0]
now = datetime.datetime.utcnow().isoformat()+'Z'
def addmov(**k):
    mid[0]+=1; base=dict(id=mid[0], da_location_id=None, a_location_id=None, stato='confermato', fornitore=None, prezzo_unitario=None, nota=None, creato_da='import', data_movimento=now, data_conferma=None, created_at=now, updated_at=now)
    base.update(k); movs.append(base); return base
for (wi, q, row, loc, pr, az) in d['m']:
    wid = wi+1
    if loc.startswith('Pending Delivery - '):
        addmov(wine_id=wid, tipo='acquisto', qty=q, a_location_id=LID['Bixio'], stato='in_corso', fornitore=loc[19:].replace(' - Primeur',''), prezzo_unitario=pr, nota='Import Cantina.xlsx riga %d'%row); continue
    dest = 'Bixio' if loc.lower()=='bixio (check)' else ('Da ritrovare' if loc=='Check' else loc)
    addmov(wine_id=wid, tipo='carico_iniziale', qty=q, a_location_id=LID[dest], prezzo_unitario=pr)
    if loc.lower()=='bixio (check)' or loc=='Check': segn.append(dict(id=len(segn)+1, wine_id=wid, location_id=LID[dest], qty=q, motivo='test '+loc, aperta=True, esito=None))
    if az=='Bixio': addmov(wine_id=wid, tipo='trasferimento', qty=q, da_location_id=LID[dest], a_location_id=LID['Bixio'], stato='pianificato')
membri=[dict(email='matteo.robin99@gmail.com', nome='Matteo')]
def giac():
    g={}
    for m in movs:
        if m['stato']=='confermato' and m['a_location_id']: g[(m['wine_id'],m['a_location_id'])]=g.get((m['wine_id'],m['a_location_id']),0)+m['qty']
        if m['stato'] in ('in_corso','confermato') and m['da_location_id']: g[(m['wine_id'],m['da_location_id'])]=g.get((m['wine_id'],m['da_location_id']),0)-m['qty']
    return [dict(wine_id=k[0], location_id=k[1], qty=v) for k,v in g.items() if v!=0]
TABLES = dict(wines=wines, locations=locs, movimenti=movs, segnalazioni=segn, membri=membri)
log=[]
def filt(rows, qs):
    out=rows
    for k,v in qs.items():
        if k in ('select','order','limit','offset','columns'): continue
        v=v[0]
        if v.startswith('eq.'): out=[r for r in out if str(r.get(k))==unquote(v[3:])]
        elif v.startswith('in.('): vals=v[4:-1].split(','); out=[r for r in out if str(r.get(k)) in vals]
    return out
def handler(route, req):
    u=urlparse(req.url); p=u.path; qs=parse_qs(u.query); m=req.method
    hdr={'Content-Type':'application/json','Access-Control-Allow-Origin':'*'}
    if m=='OPTIONS': return route.fulfill(status=200, headers={**hdr,'Access-Control-Allow-Headers':'*','Access-Control-Allow-Methods':'*'})
    if p.startswith('/auth/v1'): return route.fulfill(status=200, headers=hdr, body=json.dumps(SESSION))
    if p=='/rest/v1/rpc/is_member': return route.fulfill(status=200, headers=hdr, body='true')
    t=p.split('/')[-1]
    body=json.loads(req.post_data) if req.post_data else None
    log.append((m,t,qs,body))
    if t=='giacenze': rows=giac()
    else: rows=TABLES[t]
    if m=='GET':
        r=filt(rows,qs); return route.fulfill(status=200, headers={**hdr,'Content-Range':f'0-{len(r)}/{len(r)}'}, body=json.dumps(r))
    if m=='POST':
        items=body if isinstance(body,list) else [body]; created=[]
        for it in items:
            if t=='movimenti':
                it={k:v for k,v in it.items()}
                x=addmov(**{**dict(creato_da='matteo.robin99@gmail.com'), **it})
                if x['da_location_id'] and x['stato'] in ('in_corso','confermato'):
                    av=sum(g['qty'] for g in giac() if g['wine_id']==x['wine_id'] and g['location_id']==x['da_location_id'])
                    if av<0: movs.remove(x); return route.fulfill(status=400, headers=hdr, body=json.dumps({'message':'Giacenza insufficiente: disponibili %d'%(av+x['qty'])}))
                created.append(x)
            else:
                it=dict(it); it.setdefault('id', max([r.get('id',0) for r in rows if isinstance(r.get('id'),int)]+[0])+1); rows.append(it); created.append(it)
        single='vnd.pgrst.object' in (req.headers.get('accept') or '')
        return route.fulfill(status=201, headers=hdr, body=json.dumps(created[0] if single else created))
    if m=='PATCH':
        r=filt(rows,qs)
        for x in r: x.update(body)
        return route.fulfill(status=200, headers=hdr, body=json.dumps(r))
    if m=='DELETE':
        r=filt(rows,qs)
        for x in r: rows.remove(x)
        return route.fulfill(status=200, headers=hdr, body='[]')
    route.fulfill(status=404, body='{}')
exp=int(time.time())+3600*24
SESSION=dict(access_token='x.eyJzdWIiOiIxIn0.y', refresh_token='r', expires_in=86400, expires_at=exp, token_type='bearer', user=dict(id='u1', email='matteo.robin99@gmail.com', aud='authenticated', role='authenticated', app_metadata={}, user_metadata={}, created_at=now))
errors=[]
with sync_playwright() as pw:
    b=pw.chromium.launch(executable_path='/usr/bin/google-chrome', args=['--no-sandbox'])
    ctx=b.new_context(viewport={'width':390,'height':844}, device_scale_factor=2, accept_downloads=True)
    ctx.add_init_script(f"localStorage.setItem('sb-ddeoswpycvlfhztjlkdj-auth-token', {json.dumps(json.dumps(SESSION))});")
    ctx.route(re.compile(r'https://ddeoswpycvlfhztjlkdj\.supabase\.co/.*'), handler)
    pg=ctx.new_page()
    pg.on('console', lambda msg: errors.append(msg.text) if msg.type=='error' else None)
    pg.on('pageerror', lambda e: errors.append('PAGEERROR '+str(e)))
    pg.goto('file:///home/claude/cantina-app/dist/index.html'); pg.wait_for_timeout(1500)
    pg.screenshot(path='01_home.png')
    print('stats:', pg.inner_text('.stats'))
    pg.click('text=Cantinetta Robin'); pg.wait_for_timeout(300); pg.screenshot(path='02_loc.png')
    pg.click('.row >> nth=0'); pg.wait_for_timeout(300); pg.screenshot(path='03_wine.png', full_page=True)
    title=pg.inner_text('.hero h2'); print('wine:', title)
    pg.click('button[data-act=move]'); pg.wait_for_timeout(200)
    pg.select_option('select[name=to]', label='Bixio'); pg.screenshot(path='04_move.png')
    pg.click('#f button[type=submit]'); pg.wait_for_timeout(800)
    print('last post:', [l for l in log if l[0]=='POST'][-1][3])
    pg.click('[data-tab=arrivi]'); pg.wait_for_timeout(300); pg.screenshot(path='05_arrivi.png', full_page=True)
    pg.click('.mcard button[data-act=confirmmov] >> nth=0'); pg.wait_for_timeout(200)
    pg.fill('#f input[name=qty]', '1'); pg.click('#f button[type=submit]'); pg.wait_for_timeout(800)
    print('last patch:', [l for l in log if l[0]=='PATCH'][-1][3])
    pg.click('[data-tab=verifica]'); pg.wait_for_timeout(300); pg.screenshot(path='06_verifica.png')
    pg.click('button[data-act=resolve] >> nth=0'); pg.wait_for_timeout(200); pg.screenshot(path='07_resolve.png')
    pg.select_option('#f select[name=to]', label='Cantinetta Nonno'); pg.click('#f button[type=submit]'); pg.wait_for_timeout(800)
    print('resolve ops:', [(l[0],l[1],l[3]) for l in log[-6:] if l[0]!='GET'])
    pg.click('[data-tab=cerca]'); pg.fill('input.search','barolo'); pg.wait_for_timeout(300); pg.screenshot(path='08_cerca.png')
    print('cerca note:', pg.inner_text('main .note'))
    pg.click('[data-tab=altro]'); pg.wait_for_timeout(200)
    with pg.expect_download() as dl: pg.click('button[data-act=export]')
    dl.value.save_as('/home/claude/apptest/export.xlsx'); print('export ok')
    pg.click('button[data-act=buy]'); pg.wait_for_timeout(200); pg.click('button[data-mode=new]'); pg.wait_for_timeout(200); pg.screenshot(path='09_buy.png', full_page=True)
    b.close()
print('ERRORS:', errors)
