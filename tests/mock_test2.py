import json, re, time, datetime, sys
from urllib.parse import urlparse, parse_qs, unquote
from playwright.sync_api import sync_playwright

d = json.load(open('/home/claude/cantina/payload.json'))
LOCS = ['Cantinetta Robin','Cantinetta Nonno','Cantinetta Nonno Cucina','Corridoio Nonno','Bixio','Cantinetta Tadino','Casa Nonno Sardegna','Cantina Corridoni','Cantina Cadamosto']
locs = [dict(id=i+1, nome=n, tipo='fisica', ruolo=None, capacita=30 if n=='Cantinetta Robin' else (120 if n=='Cantinetta Nonno' else None), attiva=True, ordine=i+1, note=None) for i,n in enumerate(LOCS)]
locs.append(dict(id=10, nome='Da ritrovare', tipo='sconosciuta', ruolo='Posizione sconosciuta', capacita=None, attiva=True, ordine=99, note=None))
locs.append(dict(id=11, nome='Corridoni', tipo='gruppo', ruolo='Casa', capacita=None, attiva=True, ordine=1, note=None))
for l in locs:
    l.setdefault('parent_id', None)
    if l['nome'] in ('Cantinetta Nonno','Cantinetta Nonno Cucina','Corridoio Nonno','Cantinetta Robin'): l['parent_id']=11; l['ordine']=10+l['id']
    if l['nome']=='Bixio': l['ordine']=2
LID = {l['nome']: l['id'] for l in locs}
cols = ['produttore','vino','annata','formato','denominazione','regione','paese','tipologia','stile','vitigno','dw_inizio','dw_fine','fonte_dw','bere_presto']
wines = [dict(id=i+1, note=None, **dict(zip(cols, w))) for i,w in enumerate(d['w'])]
movs=[]; segn=[]; mid=[0]
now = datetime.datetime.utcnow().isoformat()+'Z'
def addmov(**k):
    mid[0]+=1; base=dict(id=mid[0], motivo=None, da_location_id=None, a_location_id=None, stato='confermato', fornitore=None, prezzo_unitario=None, nota=None, creato_da='import', data_movimento=now, data_conferma=None, created_at=now, updated_at=now)
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
    pg.screenshot(path='s1_home.png', full_page=True)
    print('kpi:', pg.inner_text('.kpis').replace(chr(10),' '))
    pg.click('button[data-act=comploc][data-loc="11"]'); pg.wait_for_timeout(200)
    pg.click('button[data-act=compdim][data-dim=fin]'); pg.wait_for_timeout(200)
    print('comp corridoni fin:', pg.inner_text('.bars').replace(chr(10),' | '))
    pg.click('.bars .brow >> nth=0'); pg.wait_for_timeout(300)
    print('cerca after comp:', pg.inner_text('main .note >> nth=0'))
    pg.screenshot(path='s2_cerca.png')
    pg.click('[data-act=home]'); pg.wait_for_timeout(200)
    pg.click('button[data-act=gotogli]'); pg.wait_for_timeout(300)
    pg.fill('input[data-bind=outQ]','leoville'); pg.wait_for_timeout(300)
    pg.screenshot(path='s3_togli.png')
    pg.click('main .lrow[data-act=out] >> nth=0'); pg.wait_for_timeout(200)
    pg.click('fieldset.seg-r label:has-text("Regalata")'); pg.screenshot(path='s4_out.png')
    pg.click('#f button[type=submit]'); pg.wait_for_timeout(800)
    print('out post:', [l for l in log if l[0]=='POST'][-1][3])
    pg.click('button[data-act=goaggiungi]'); pg.wait_for_timeout(200); pg.screenshot(path='s5_aggiungi.png')
    pg.click('button[data-act=found]'); pg.wait_for_timeout(200)
    pg.fill('#buyq','sassicaia'); pg.wait_for_timeout(200)
    pg.click('.sheet .lrow >> nth=0'); pg.wait_for_timeout(200)
    pg.select_option('#f select[name=to]', label='Cantinetta Robin'); pg.screenshot(path='s6_found.png')
    pg.click('#f button[type=submit]'); pg.wait_for_timeout(800)
    print('found post:', [l for l in log if l[0]=='POST'][-1][3])
    pg.click('button.morebtn'); pg.wait_for_timeout(200); pg.screenshot(path='s7_altro.png')
    pg.click('button[data-act=goluoghi]'); pg.wait_for_timeout(200); pg.screenshot(path='s8_luoghi.png')
    pg.click('[data-act=home]'); pg.wait_for_timeout(200)
    pg.click('.lrow.top[data-act=openloc] >> nth=0'); pg.wait_for_timeout(300); pg.screenshot(path='s9_corridoni.png')
    pg.click('main .lrow[data-act=openwine] >> nth=0'); pg.wait_for_timeout(300)
    pg.click('button[data-act=move]'); pg.wait_for_timeout(200); pg.screenshot(path='s10_move.png')
    pg.select_option('#f select[name=to]', label='Bixio'); pg.click('#f button[type=submit]'); pg.wait_for_timeout(800)
    print('move post:', [l for l in log if l[0]=='POST'][-1][3])
    pg.click('button.morebtn'); pg.wait_for_timeout(200)
    with pg.expect_download() as dl: pg.click('button[data-act=export]')
    dl.value.save_as('/home/claude/apptest/export2.xlsx'); print('export ok')
    b.close()
print('ERRORS:', errors)
