import streamlit as st
import pandas as pd
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import hashlib
import re
import json
import time
import math
import io
import zipfile
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from supabase import create_client

st.set_page_config(page_title="Ascensori Servizi & Tesyx", page_icon="🛠️", layout="wide")
st.markdown("""
<style>
@media (max-width: 700px) {
  .block-container {padding-top: 1rem; padding-left: .65rem; padding-right: .65rem;}
  div[data-testid="stMetric"] {border: 1px solid rgba(128,128,128,.25); padding: .45rem; border-radius: .65rem;}
  .stButton > button {min-height: 3rem; font-weight: 700;}
}
  div[class*="st-key-registration_button_"] button {background-color:#15803d !important; border-color:#15803d !important; color:white !important;}
  div[class*="st-key-blocked_"] button {background-color:#b91c1c !important; border-color:#b91c1c !important; color:white !important; opacity:1 !important;}
</style>
""", unsafe_allow_html=True)
MESI=[('Gen','Gennaio',1),('feb','Febbraio',2),('mar','Marzo',3),('apr','Aprile',4),('mag','Maggio',5),('giu','Giugno',6),('lug','Luglio',7),('ago','Agosto',8),('sett','Settembre',9),('ott','Ottobre',10),('nov','Novembre',11),('dic','Dicembre',12)]
XLS='impianti.xlsx'

# Nuovo archivio indipendente dal precedente
TABLES = {'impianti':'impianti_v2','programmazione':'programmazione_v2','manutenzioni':'manutenzioni_v2','storico_annullamenti':'storico_annullamenti_v2','tecnici':'tecnici_v2'}

def table(name):
    return TABLES.get(name, name)


@st.cache_resource
def db():
    url = st.secrets.get("SUPABASE_URL")
    key = st.secrets.get("SUPABASE_KEY") or st.secrets.get("SUPABASE_SERVICE_KEY")
    if not url or not key:
        raise RuntimeError("Mancano SUPABASE_URL o SUPABASE_KEY nei Secrets di Streamlit.")
    if str(url).rstrip("/") != "https://ysbcbafaoubeojcnpmcc.supabase.co":
        raise RuntimeError("BLOCCO DI SICUREZZA: questa versione può collegarsi solo al Supabase di prova.")
    return create_client(url, key)

def clean(v):
    if pd.isna(v): return ''
    return str(v).strip()


def months_for_cadence(code):
    code=clean(code).upper().replace(' ', '')
    if not code: return []  # impianto presente, cadenza da assegnare
    if code in ('M','M1','MENSILE'): return list(range(1,13))
    if code == 'ST8': return [8]
    if code == 'ST10': return [10]
    if code in ('ST8-3/10', 'ST8-03/10'): return [3,10]
    match=re.fullmatch(r'([BTQS])([1-9][0-9]*)',code)
    if not match: return None
    period={'B':2,'T':3,'Q':4,'S':6}[match.group(1)]
    offset=(int(match.group(2))-1)%period+1
    return list(range(offset,13,period))


def import_frame(upload):
    """Legge esclusivamente l'archivio unificato, preservando S/N/? senza deduzioni."""
    df = pd.read_excel(upload, sheet_name='Archivio unificato', dtype=str).fillna('')
    df.columns = [str(c).strip() for c in df.columns]
    required = ['Cod. cliente','Cliente','Matricola','Num. fabbrica','Comune','Indirizzo',
                'Provincia','Tecnico proposto','Cadenza Roberto','Esito controllo','Note verifica',
                'Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError('Archivio non compatibile. Colonne mancanti: '+', '.join(missing))
    month_excel = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic']
    for c in month_excel:
        df[c] = df[c].astype(str).str.strip().str.upper()
    invalid = [(i+2,c,df.at[i,c]) for i in df.index for c in month_excel
               if df.at[i,c] not in ('S','N','?')]
    if invalid:
        raise ValueError(f'{len(invalid)} celle mensili non valide (prime: {invalid[:5]}). Sono ammessi solo S, N e ?.')
    if df.empty:
        raise ValueError('Archivio vuoto.')
    # Identificativi distinti per ogni riga; l'archivio deve essere importato una sola volta.
    df['codice_import'] = ['UNI-2026-'+str(i+1).zfill(5) for i in range(len(df))]
    return df


def import_new_archive(df):
    """Importazione nel SOLO database di prova; ? salvato come NULL, mai come N."""
    client = db()  # il controllo URL blocca ogni altro progetto
    count = client.table('impianti_v2').select('id',count='exact').limit(1).execute().count or 0
    if count:
        raise ValueError('Archivio già popolato: importazione bloccata. Non vengono cancellati dati.')
    month_excel = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic']
    month_db = [name.lower() for name,_,_ in MESI]
    records=[]
    for _,r in df.iterrows():
        def val(col): return clean(r.get(col)) or None
        prog=val('Riga Roberto')
        try: prog=int(float(prog)) if prog else None
        except ValueError: prog=None
        # Le note di verifica e gli esiti sono conservati senza nascondere anomalie.
        notes=[f"ESITO VERIFICA: {val('Esito controllo') or '-'}",
               f"NOTE VERIFICA: {val('Note verifica') or '-'}",
               f"NOTE ROBERTO: {val('Note Roberto') or '-'}",
               f"ORIGINE: {val('Origine') or '-'}",
               f"CODICE DEFINITIVO: {val('Codice Definitivo') or '-'}"]
        records.append({'codice':r['codice_import'],'codice_cliente':val('Cod. cliente'),
          'cliente':val('Cliente'),'matricola':val('Matricola'),'numero_fabbrica':val('Num. fabbrica'),
          'cap':None,'comune':val('Comune'),'indirizzo':val('Indirizzo'),'provincia':val('Provincia'),
          'tipo_manutenzione':val('Cadenza Roberto'),'progressivo':prog,
          'tecnico_assegnato':val('Tecnico proposto') or 'Non assegnato',
          'note':' | '.join(notes)})
    # ATTENZIONE: inserimento in lotti, non transazionale. In caso di errore fermarsi.
    for i in range(0,len(records),150):
        client.table('impianti_v2').insert(records[i:i+150]).execute()
    invalidate_data()
    id_map={r['codice']:r['id'] for r in all_rows('impianti','id,codice')}
    plans=[]; jobs=[]
    for (_,r),rec in zip(df.iterrows(),records):
        iid=id_map[rec['codice']]
        plan={'impianto_id':iid,'anno':2026}
        for excel_col,db_col,num in zip(month_excel,month_db,range(1,13)):
            v=r[excel_col]
            plan[db_col]=True if v=='S' else False if v=='N' else None
            if v=='S':
                jobs.append({'impianto_id':iid,'anno_competenza':2026,'mese_competenza':num,
                             'tecnico_assegnato':rec['tecnico_assegnato'],'stato':'DA_FARE'})
        plans.append(plan)
    for i in range(0,len(plans),150):
        client.table('programmazione_v2').insert(plans[i:i+150]).execute()
    for i in range(0,len(jobs),150):
        client.table('manutenzioni_v2').insert(jobs[i:i+150]).execute()
    names=sorted(set(r['tecnico_assegnato'] for r in records if r['tecnico_assegnato']!='Non assegnato'))
    for name in names:
        client.table('tecnici_v2').upsert({'nome':name,'attivo':True},on_conflict='nome').execute()
    invalidate_data()
    return len(records),len(jobs),len(names)


def recover_partial_import(df):
    """Complete ONLY matching, unfinished 2026 test import; never overwrite existing rows."""
    client=db()
    expected_codes=set(df['codice_import'].tolist())
    plants=fetch_rows('impianti','id,codice,tecnico_assegnato')
    mapping={r['codice']:r for r in plants}
    if len(plants)!=len(df) or set(mapping)!=expected_codes:
        raise ValueError('Gli impianti nel database non corrispondono esattamente al file: recupero bloccato.')
    plans_existing=fetch_rows('programmazione','*')
    jobs_existing=fetch_rows('manutenzioni','impianto_id,anno_competenza,mese_competenza,stato')
    if any(int(r['anno'])!=2026 for r in plans_existing) or any(int(r['anno_competenza'])!=2026 for r in jobs_existing):
        raise ValueError('Sono presenti altri anni: recupero automatico bloccato.')
    if any(r['stato']!='DA_FARE' for r in jobs_existing):
        raise ValueError('Sono presenti manutenzioni modificate o eseguite: recupero bloccato.')
    months=['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic']
    keys=['gen','feb','mar','apr','mag','giu','lug','ago','sett','ott','nov','dic']
    plans=[]; jobs=[]
    for _,row in df.iterrows():
        plant=mapping[row['codice_import']]
        iid=int(plant['id'])
        plan={'impianto_id':iid,'anno':2026}
        for col,key,num in zip(months,keys,range(1,13)):
            v=row[col]
            plan[key]=True if v=='S' else False if v=='N' else None
            if v=='S':
                jobs.append({'impianto_id':iid,'anno_competenza':2026,'mese_competenza':num,'tecnico_assegnato':plant['tecnico_assegnato'],'stato':'DA_FARE'})
        plans.append(plan)
    plan_by_id={int(r['impianto_id']):r for r in plans}
    if len(plans_existing)!=len({int(r['impianto_id']) for r in plans_existing}):
        raise ValueError('Programmazioni duplicate: recupero bloccato.')
    for old in plans_existing:
        desired=plan_by_id.get(int(old['impianto_id']))
        if desired is None or any(old.get(k)!=desired[k] for k in keys):
            raise ValueError('Programmazioni già presenti diverse dal file: recupero bloccato.')
    job_keys={(int(j['impianto_id']),int(j['anno_competenza']),int(j['mese_competenza'])) for j in jobs}
    for old in jobs_existing:
        if (int(old['impianto_id']),int(old['anno_competenza']),int(old['mese_competenza'])) not in job_keys:
            raise ValueError('Manutenzioni esistenti non previste dal file: recupero bloccato.')
    existing_plan_ids={int(r['impianto_id']) for r in plans_existing}
    existing_jobs={(int(r['impianto_id']),int(r['anno_competenza']),int(r['mese_competenza'])) for r in jobs_existing}
    missing_plans=[r for r in plans if r['impianto_id'] not in existing_plan_ids]
    missing_jobs=[r for r in jobs if (r['impianto_id'],r['anno_competenza'],r['mese_competenza']) not in existing_jobs]
    for i in range(0,len(missing_plans),100):
        client.table('programmazione_v2').insert(missing_plans[i:i+100]).execute()
    for i in range(0,len(missing_jobs),100):
        client.table('manutenzioni_v2').insert(missing_jobs[i:i+100]).execute()
    techs=sorted({r['tecnico_assegnato'] for r in plants if r.get('tecnico_assegnato') and r['tecnico_assegnato']!='Non assegnato'})
    for name in techs:
        client.table('tecnici_v2').upsert({'nome':name,'attivo':True},on_conflict='nome').execute()
    invalidate_data()
    return len(missing_plans),len(missing_jobs),len(techs)


@st.cache_data(ttl=45, show_spinner=False)
def cached_rows(table_name, cols='*', filters=()):
    # Cache letture; non memorizza credenziali o connessioni.
    return fetch_rows(table_name, cols, filters)

def fetch_rows(table_name, cols='*', filters=()):
    s=db(); rows=[]; start=0; step=1000
    while True:
        query=s.table(globals()['table'](table_name)).select(cols)
        for column, value in filters:
            query=query.eq(column, value)
        r=query.range(start,start+step-1).execute().data
        rows.extend(r)
        if len(r)<step: break
        start+=step
    return rows

@st.cache_data(ttl=20, show_spinner=False)
def official_active():
    rows=fetch_rows('impianti','id,codice,note',(('codice','OFF-2026-00001'),))
    return bool(rows and '[ARCHIVIO_UFFICIALE_ATTIVO]' in (rows[0].get('note') or ''))

def all_rows(table, cols='*'):
    rows=cached_rows(table, cols)
    if table=='impianti' and official_active():
        return [r for r in rows if str(r.get('codice','')).startswith('OFF-2026-')]
    return rows

def invalidate_data():
    cached_rows.clear()
    official_active.clear()
    get_status_overrides.clear()
    dataset.clear()

def seed_if_needed():
    # La nuova base dati viene caricata esclusivamente da Amministratore.
    return

def ensure_jobs(year):
    s=db(); imps=pd.DataFrame(all_rows('impianti','id,codice,matricola,cliente,indirizzo,comune,provincia,tecnico_assegnato'))
    if imps.empty: return
    prs=pd.DataFrame(all_rows('programmazione','*'))
    if not prs.empty and year not in set(prs.get('anno',pd.Series(dtype=int)).tolist()):
        # for test, copy annual plan from 2026 to selected year
        base=prs[prs.anno==min(prs.anno)].copy()
        if not base.empty:
            vals=[]
            for _,r in base.iterrows():
                x={'impianto_id':int(r.impianto_id),'anno':year}
                for name,_,_ in MESI: x[name.lower()]=r[name.lower()]
                vals.append(x)
            for i in range(0,len(vals),200): s.table('programmazione_v2').insert(vals[i:i+200]).execute()
            prs=pd.DataFrame(all_rows('programmazione','*'))
    existing={(x['impianto_id'],x['anno_competenza'],x['mese_competenza']) for x in all_rows('manutenzioni','impianto_id,anno_competenza,mese_competenza')}
    vals=[]
    if prs.empty: return
    for _,r in prs[prs.anno==year].iterrows():
        tech=imps.loc[imps.id==r.impianto_id,'tecnico_assegnato']; tech=tech.iloc[0] if len(tech) else None
        for name,_,num in MESI:
            if r[name.lower()] is True and (int(r.impianto_id),year,num) not in existing:
                vals.append({'impianto_id':int(r.impianto_id),'anno_competenza':year,'mese_competenza':num,'tecnico_assegnato':tech,'stato':'DA_FARE'})
    for i in range(0,len(vals),200): s.table('manutenzioni_v2').insert(vals[i:i+200]).execute()

def generate_2027_from_2026():
    """Ripresa controllata 2027 sul solo Supabase di prova.

    Mantiene S/N/? (True/False/None), non aggiorna mai record esistenti,
    non modifica le manutenzioni 2026 e può riprendere lotti interrotti.
    """
    client = db()  # Controllo vincolante sull'URL del progetto di prova.
    keys = ['gen','feb','mar','apr','mag','giu','lug','ago','sett','ott','nov','dic']
    plants = fetch_rows('impianti', 'id,tecnico_assegnato')
    source = fetch_rows('programmazione', 'impianto_id,anno,' + ','.join(keys), (('anno',2026),))
    target = fetch_rows('programmazione', 'impianto_id,anno,' + ','.join(keys), (('anno',2027),))
    jobs_2027 = fetch_rows('manutenzioni', 'impianto_id,anno_competenza,mese_competenza,stato', (('anno_competenza',2027),))
    plant_ids = {int(p['id']) for p in plants}
    if not plant_ids or len(plant_ids) != 2547:
        raise ValueError(f'Attesi 2547 impianti di prova; trovati {len(plant_ids)}. Operazione bloccata.')
    source_by_id = {int(r['impianto_id']):r for r in source}
    if len(source_by_id) != len(source) or set(source_by_id) != plant_ids:
        raise ValueError('Programmazione 2026 incompleta, duplicata o non corrispondente agli impianti: operazione bloccata.')
    target_by_id = {int(r['impianto_id']):r for r in target}
    if len(target_by_id) != len(target) or not set(target_by_id).issubset(plant_ids):
        raise ValueError('Programmazioni 2027 duplicate o sconosciute: operazione bloccata.')
    for iid, old in target_by_id.items():
        if any(old.get(k) != source_by_id[iid].get(k) for k in keys):
            raise ValueError(f'Programmazione 2027 già modificata per impianto {iid}: nessuna sovrascrittura consentita.')
    expected_jobs = {(iid, month) for iid,r in source_by_id.items()
                     for month,k in enumerate(keys,1) if r.get(k) is True}
    existing_jobs = set()
    for job in jobs_2027:
        pair = (int(job['impianto_id']),int(job['mese_competenza']))
        if pair in existing_jobs or pair not in expected_jobs or job['stato'] != 'DA_FARE':
            raise ValueError('Manutenzioni 2027 già modificate, duplicate o non previste: operazione bloccata.')
        existing_jobs.add(pair)
    missing_plans = [{'impianto_id':iid,'anno':2027,**{k:r.get(k) for k in keys}}
                     for iid,r in source_by_id.items() if iid not in target_by_id]
    tech_by_id = {int(p['id']):p.get('tecnico_assegnato') for p in plants}
    missing_jobs = [{'impianto_id':iid,'anno_competenza':2027,'mese_competenza':month,
                     'tecnico_assegnato':tech_by_id[iid],'stato':'DA_FARE'}
                    for iid,month in sorted(expected_jobs - existing_jobs)]
    # Preflight completato: solo INSERT, mai UPDATE/DELETE.
    for start in range(0,len(missing_plans),100):
        client.table('programmazione_v2').insert(missing_plans[start:start+100]).execute()
    for start in range(0,len(missing_jobs),100):
        client.table('manutenzioni_v2').insert(missing_jobs[start:start+100]).execute()
    invalidate_data()
    return len(missing_plans),len(missing_jobs),len(expected_jobs)


@st.cache_data(ttl=15)
def dataset(year):
    jobs=pd.DataFrame(cached_rows('manutenzioni','id,impianto_id,anno_competenza,mese_competenza,tecnico_assegnato,tecnico_esecutore,stato,eseguita_il,semestrale,biennale,scadenza_biennale,note',(('anno_competenza',year),)))
    imps=pd.DataFrame(all_rows('impianti','id,codice,matricola,cliente,indirizzo,comune,provincia,tecnico_assegnato,progressivo,note'))
    if jobs.empty or imps.empty: return pd.DataFrame()
    p=jobs[jobs.anno_competenza==year].merge(imps,left_on='impianto_id',right_on='id',suffixes=('','_imp'))
    p['mese_nome']=p.mese_competenza.map({n:nome for _,nome,n in MESI})
    now=datetime.now(); current=now.month if year==now.year else (13 if year<now.year else 0)
    p['assegnato']=p['tecnico_assegnato_imp'].fillna(p['tecnico_assegnato']).fillna('Non assegnato')
    p=p.rename(columns={'note_imp':'note_impianto'})
    return apply_status(p)

# Lo stato e le date di riattivazione sono in una tabella separata: nessuna
# modifica ai 2.351 impianti, alle programmazioni o allo storico esistente.
STATUS_TABLE = 'stati_impianti_v2'

@st.cache_data(ttl=45, show_spinner=False)
def get_status_overrides():
    return {int(x['impianto_id']): x for x in all_rows(STATUS_TABLE, '*')}

def initial_status(note):
    text = clean(note).upper()
    # Solo indicazioni esplicite, evitando frasi come "non fermo".
    if re.search(r'\b(NON\s+FERMO|RIATTIVAT[OA]|RIMESSO\s+IN\s+SERVIZIO)\b', text):
        return 'ATTIVO'
    if re.search(r'\b(SOSPES[OA]|SOSPENSIONE)\b', text):
        return 'SOSPESO'
    if re.search(r'\b(FERM[OA]|FUORI\s+SERVIZIO)\b', text):
        return 'FERMO'
    if re.search(r'\b(DISDETT[OA])\b', text):
        return 'DISDETTO'
    if re.search(r'\b(CHIUS[OA])\b', text):
        return 'CHIUSO'
    return 'ATTIVO'

def start_month_for_tech(name):
    # Solo Cristian inizia a ottobre 2026; tutti gli altri a novembre.
    return 10 if re.search(r'\bCRISTIAN\b', clean(name).upper()) else 11

def apply_status(jobs):
    if jobs.empty: return jobs
    overrides = get_status_overrides()
    jobs = jobs.copy()
    def row_status(r):
        override = overrides.get(int(r['impianto_id']))
        status = override['stato'] if override else initial_status(r.get('note_impianto'))
        if r['stato'] == 'ESEGUITA': return 'ESEGUITA'
        if status != 'ATTIVO': return 'DA NON FARE'
        # In caso di riattivazione non si recuperano i mesi antecedenti.
        month = int(r['mese_competenza']); year = int(r['anno_competenza'])
        if (year, month) < (2026, start_month_for_tech(r.get('assegnato'))):
            return 'NON AVVIATA'
        if override and override.get('data_riattivazione'):
            try:
                d = date.fromisoformat(str(override['data_riattivazione'])[:10])
                if (year,month) < (d.year,d.month): return 'NON AVVIATA'
            except ValueError: pass
        now = date.today()
        if (year,month) < (now.year,now.month): return 'SCADUTA'
        return 'DA FARE'
    jobs['stato_impianto'] = jobs['impianto_id'].map(lambda x: overrides.get(int(x), {}).get('stato'))
    jobs['stato_impianto'] = jobs.apply(lambda r: r['stato_impianto'] or initial_status(r.get('note_impianto')),axis=1)
    jobs['stato_ui'] = jobs.apply(row_status,axis=1)
    return jobs

def save_plant_status(plant_id, status, reason):
    old = get_status_overrides().get(int(plant_id))
    today = date.today().isoformat()
    payload = {'impianto_id':int(plant_id),'stato':status, 'data_stato':today, 'aggiornato_il':datetime.now().astimezone().isoformat(),
               'motivo':reason.strip() or None,
               'data_riattivazione': today if status=='ATTIVO' and (not old or old['stato']!='ATTIVO') else (old.get('data_riattivazione') if old else None)}
    db().table(STATUS_TABLE).upsert(payload,on_conflict='impianto_id').execute()
    invalidate_data()


def street_sort_parts(address):
    """Raggruppa le varianti di una via e ordina i civici numericamente.

    Non stima distanze: per un percorso di prossimita servono coordinate.
    """
    raw=clean(address).upper()
    raw=re.sub(r"[.,;]", " ", raw)
    raw=re.sub(r"\s+", " ", raw).strip()
    raw=re.sub(r"^(VIALE|VIA|VLE|V\s+LE|V\s+|PIAZZA|PZZA|P\s+ZA|PIAZZALE|LARGO|CORSO|STRADA|LOC\s+|LOCALITA)\s+", "", raw)
    # Il civico viene estratto dal termine finale (es. 28/C, 28 C/O, 28).
    found=re.search(r"(?:^|\s)(\d{1,4})(?:\s*[/\-]?\s*([A-Z]))?(?:\s|$)", raw)
    if found:
        street=raw[:found.start()].strip()
        civic=int(found.group(1))
        suffix=found.group(2) or ''
    else:
        street=raw
        civic=999999
        suffix=''
    return street or raw, civic, suffix


def operators():
    x=all_rows('tecnici','nome,attivo'); return sorted([r['nome'] for r in x if r.get('attivo',True)])

def close_job(job_id, oper, dt, sem, biennale, note):
    # Controllo sulla manutenzione effettiva, senza limitare l'operativita al 2027.
    rows = db().table('manutenzioni_v2').select('anno_competenza,mese_competenza').eq('id', int(job_id)).execute().data
    if len(rows) != 1:
        raise ValueError('Manutenzione non trovata o duplicata.')
    year = int(rows[0]['anno_competenza'])
    month = int(rows[0]['mese_competenza'])
    if year not in (2026, 2027):
        raise ValueError('Anno di competenza non abilitato.')
    current = dataset(year)
    match = current[current['id'] == int(job_id)] if not current.empty else pd.DataFrame()
    if match.empty or match.iloc[0]['stato_ui'] not in ('DA FARE', 'SCADUTA'):
        raise ValueError('Manutenzione non registrabile: impianto sospeso, non avviato o gia eseguito.')
    row = match.iloc[0]
    first_month = start_month_for_tech(row['assegnato']) if year == 2026 else 1
    first_day = date(year, first_month, 1)
    if dt < first_day or dt > date.today():
        raise ValueError(f'La data di esecuzione deve essere compresa tra il {first_day:%d/%m/%Y} e oggi.')
    if (year, month) < (2026, start_month_for_tech(row['assegnato'])):
        raise ValueError('Manutenzione precedente alla data di avvio del tecnico.')
    ts = datetime.combine(dt, datetime.now().time()).astimezone().isoformat()
    scadenza = (dt + relativedelta(months=24)).isoformat() if biennale else None
    db().table('manutenzioni_v2').update({'stato':'ESEGUITA','tecnico_esecutore':oper,'eseguita_il':ts,'semestrale':bool(sem),'biennale':bool(biennale),'scadenza_biennale':scadenza,'note':note}).eq('id',int(job_id)).eq('stato','DA_FARE').execute()
    invalidate_data()

def cancel_job(job_id, reason):
    # Conserva una copia completa PRIMA di modificare il record.
    client = db()
    result = client.table('manutenzioni_v2').select('*').eq('id',int(job_id)).execute().data
    if len(result)!=1 or result[0].get('stato')!='ESEGUITA':
        raise ValueError('Questa manutenzione non risulta eseguita o non esiste più.')
    old = result[0]
    audit = {
        'manutenzione_id': int(job_id),
        'dati_precedenti': old,
        'motivo': reason,
        'annullato_il': datetime.now().astimezone().isoformat()
    }
    client.table('storico_annullamenti_v2').insert(audit).execute()
    client.table('manutenzioni_v2').update({
        'stato':'DA_FARE', 'tecnico_esecutore':None,
        'eseguita_il':None, 'semestrale':False, 'biennale':False, 'scadenza_biennale':None, 'note':None
    }).eq('id',int(job_id)).eq('stato','ESEGUITA').execute()
    invalidate_data()

def update_plan_safely(plant_id, year, months, reason):
    """Edit annual S/N plan; preserve completed jobs; log change before mutation."""
    client=db()
    current=client.table('programmazione_v2').select('*').eq('impianto_id',int(plant_id)).eq('anno',int(year)).execute().data
    if len(current)!=1:
        raise ValueError('Programmazione annuale mancante o duplicata: nessuna modifica effettuata.')
    before=current[0]
    changed={key:value for key,value in months.items() if before.get(key)!=value}
    if not changed:
        return 'Nessuna modifica: la programmazione è già aggiornata.'
    existing=client.table('manutenzioni_v2').select('id,mese_competenza,stato').eq('impianto_id',int(plant_id)).eq('anno_competenza',int(year)).execute().data
    by_month={}
    for job in existing:
        by_month.setdefault(int(job['mese_competenza']),[]).append(job)
    for name,_,num in MESI:
        if name.lower() in changed and months[name.lower()] is not True:
            if any(j['stato']=='ESEGUITA' for j in by_month.get(num,[])):
                raise ValueError(f'{name}: manutenzione già eseguita. Non è possibile impostare N.')
    # Require an audit table before any data mutation.
    # Supabase permissions must allow this read and write; do not skip audit.
    try:
        client.table('storico_programmazione_v2').select('id').limit(1).execute()
    except Exception as exc:
        raise ValueError('Manca la tabella storico_programmazione_v2: configurarla prima di abilitare il salvataggio.') from exc
    audit={'impianto_id':int(plant_id),'anno':int(year),'prima':{k:before.get(k) for k in months},'dopo':dict(months),'motivo':reason,'modificato_il':datetime.now().astimezone().isoformat()}
    # Record intent first: an interrupted operation can be reconciled from audit.
    client.table('storico_programmazione_v2').insert(audit).execute()
    client.table('programmazione_v2').update(dict(months)).eq('id',before['id']).execute()
    created=0; removed=0
    for name,_,num in MESI:
        key=name.lower()
        if key not in changed: continue
        jobs=by_month.get(num,[])
        if months[key] is True:
            if not jobs:
                plant=client.table('impianti_v2').select('tecnico_assegnato').eq('id',int(plant_id)).single().execute().data
                client.table('manutenzioni_v2').insert({'impianto_id':int(plant_id),'anno_competenza':int(year),'mese_competenza':num,'tecnico_assegnato':plant.get('tecnico_assegnato'),'stato':'DA_FARE'}).execute()
                created+=1
        else:
            for job in jobs:
                if job['stato']!='ESEGUITA':
                    client.table('manutenzioni_v2').delete().eq('id',job['id']).neq('stato','ESEGUITA').execute()
                    removed+=1
    invalidate_data()
    return f'Programmazione aggiornata. Manutenzioni aggiunte: {created}; non eseguite rimosse: {removed}.'

try:
    seed_if_needed()
except Exception as e:
    st.error('Impossibile inizializzare il database. Controlla i Secrets di Streamlit.'); st.exception(e); st.stop()

# Geocodifica solo su richiesta. Nessuna modifica a Supabase.
@st.cache_data(ttl=60*60*24*30, show_spinner=False)
def geocode_address(indirizzo, comune, provincia):
    address = ', '.join(x for x in [clean(indirizzo), clean(comune), clean(provincia), 'Italia'] if x)
    url = 'https://nominatim.openstreetmap.org/search?' + urlencode({'q': address, 'format': 'json', 'limit': 1, 'countrycodes': 'it'})
    req = Request(url, headers={'User-Agent': 'TesyxRoutePlanner/1.0 (internal maintenance planning)', 'Accept': 'application/json'})
    with urlopen(req, timeout=12) as response:
        results = json.load(response)
    if not results:
        return None
    return (float(results[0]['lat']), float(results[0]['lon']))

def geo_distance(a, b):
    lat1, lon1, lat2, lon2 = map(math.radians, [a[0], a[1], b[0], b[1]])
    h = math.sin((lat2-lat1)/2)**2 + math.cos(lat1)*math.cos(lat2)*math.sin((lon2-lon1)/2)**2
    return 6371 * 2 * math.asin(min(1, math.sqrt(h)))

def geographic_order(frame, coordinates):
    # Raggruppa per stabile: tutte le manutenzioni allo stesso indirizzo restano unite.
    groups = {}
    for _, row in frame.iterrows():
        key = (clean(row['comune']).upper(), clean(row['indirizzo']).upper(), clean(row.get('provincia')).upper())
        groups.setdefault(key, []).append(row['id'])
    valid = [key for key in groups if coordinates.get(key)]
    missing = [key for key in groups if not coordinates.get(key)]
    if not valid:
        return frame, len(missing)
    # Prova partenze diverse per evitare di fissare arbitrariamente un punto di partenza.
    starts = valid if len(valid) <= 60 else valid[:60]
    best = None
    best_length = float('inf')
    for start in starts:
        remaining = set(valid)
        remaining.remove(start)
        route = [start]
        length = 0.0
        while remaining:
            current = route[-1]
            nxt = min(remaining, key=lambda k: geo_distance(coordinates[current], coordinates[k]))
            length += geo_distance(coordinates[current], coordinates[nxt])
            route.append(nxt)
            remaining.remove(nxt)
        if length < best_length:
            best, best_length = route, length
    order = {job_id: index for index, key in enumerate(best + sorted(missing)) for job_id in groups[key]}
    return frame.assign(_giro=frame['id'].map(order)).sort_values('_giro', kind='stable').drop(columns=['_giro']), len(missing)

def activate_official_archive(source):
    """Importazione aggiuntiva, riprendibile. Nessun UPDATE/DELETE delle righe storiche.
    Attivazione finale soltanto dopo controllo impianti, piani e manutenzioni.
    """
    client=db()
    src=source.reset_index(drop=True)
    if len(src)!=2399:
        raise ValueError(f'Attesi 2399 impianti, trovati {len(src)}. Nessuna importazione.')
    keys=['gen','feb','mar','apr','mag','giu','lug','ago','sett','ott','nov','dic']
    records=[]; plans=[]; jobs=[]
    for i,r in src.iterrows():
        code='OFF-2026-'+str(i+1).zfill(5)
        cadence=clean(r['TIPO MANUT.'])
        months=months_for_cadence(cadence)
        if months is None:
            months=[] # cadenza da definire manualmente, mai dedotta
        tech=clean(r['MANUTENTORE']) or 'Non assegnato'
        records.append({'codice':code,'codice_cliente':clean(r['COD. CLIENTE']) or None,
            'cliente':clean(r['RAG. SOCIALE']) or None,'matricola':clean(r['NUMERO MATRICOLA']) or None,
            'numero_fabbrica':clean(r['NUMERO FABBRICA']) or None,
            'comune':clean(r["CITTA'"]) or None,'indirizzo':clean(r['UBICAZIONE']) or None,
            'provincia':clean(r['PROV.']) or None,'tipo_manutenzione':cadence or None,
            'tecnico_assegnato':tech,'note':'ARCHIVIO UFFICIALE 2026 | CADENZA DA DEFINIRE' if not cadence else 'ARCHIVIO UFFICIALE 2026'})
    current=fetch_rows('impianti','id,codice,note')
    official={r['codice']:r for r in current if str(r.get('codice','')).startswith('OFF-2026-')}
    allowed={r['codice'] for r in records}
    if set(official)-allowed:
        raise ValueError('Trovati codici ufficiali non previsti. Importazione bloccata.')
    for i in range(0,len(records),100):
        missing=[r for r in records[i:i+100] if r['codice'] not in official]
        if missing:
            client.table('impianti_v2').insert(missing).execute()
    current=fetch_rows('impianti','id,codice,note')
    official={r['codice']:r for r in current if str(r.get('codice','')).startswith('OFF-2026-')}
    if len(official)!=2399: raise ValueError('Numero impianti ufficiali non completo: attivazione sospesa.')
    for i,r in src.iterrows():
        iid=official['OFF-2026-'+str(i+1).zfill(5)]['id']
        months=months_for_cadence(clean(r['TIPO MANUT.'])) or []
        tech=clean(r['MANUTENTORE']) or 'Non assegnato'
        for year in (2026,2027):
            plans.append({'impianto_id':iid,'anno':year,**{k:(n in months) if clean(r['TIPO MANUT.']) else None for n,k in enumerate(keys,1)}})
            for m in months:
                jobs.append({'impianto_id':iid,'anno_competenza':year,'mese_competenza':m,'tecnico_assegnato':tech,'stato':'DA_FARE'})
    ids={v['id'] for v in official.values()}
    oldplans={(r['impianto_id'],r['anno']) for r in fetch_rows('programmazione','impianto_id,anno') if r['impianto_id'] in ids}
    for i in range(0,len(plans),100):
        missing=[r for r in plans[i:i+100] if (r['impianto_id'],r['anno']) not in oldplans]
        if missing:client.table('programmazione_v2').insert(missing).execute()
    oldjobs={(r['impianto_id'],r['anno_competenza'],r['mese_competenza']) for r in fetch_rows('manutenzioni','impianto_id,anno_competenza,mese_competenza') if r['impianto_id'] in ids}
    for i in range(0,len(jobs),100):
        missing=[r for r in jobs[i:i+100] if (r['impianto_id'],r['anno_competenza'],r['mese_competenza']) not in oldjobs]
        if missing:client.table('manutenzioni_v2').insert(missing).execute()
    allplans=fetch_rows('programmazione','impianto_id,anno')
    alljobs=fetch_rows('manutenzioni','impianto_id,anno_competenza,mese_competenza')
    plan_keys={(r['impianto_id'],r['anno']) for r in allplans if r['impianto_id'] in ids}
    job_keys={(r['impianto_id'],r['anno_competenza'],r['mese_competenza']) for r in alljobs if r['impianto_id'] in ids}
    if not {(r['impianto_id'],r['anno']) for r in plans}.issubset(plan_keys) or not {(r['impianto_id'],r['anno_competenza'],r['mese_competenza']) for r in jobs}.issubset(job_keys):
        raise ValueError('Verifica finale incompleta: vecchio archivio ancora operativo.')
    # Il marcatore attiva l’archivio solo dopo il completamento di tutti i lotti.
    first=official['OFF-2026-00001']
    client.table('impianti_v2').update({'note':(first.get('note') or '')+' [ARCHIVIO_UFFICIALE_ATTIVO]'}).eq('id',first['id']).execute()
    invalidate_data()
    return {'impianti_ufficiali':len(official),'programmazioni_ufficiali':len(plans),'manutenzioni_da_fare_ufficiali':len(jobs),'vecchio_archivio':'conservato, non cancellato'}

st.title('🛠️ Ascensori Servizi & Tesyx')
area=st.sidebar.radio('Area',['👷 Tecnico','📊 Amministratore'], index=0)
year=st.sidebar.selectbox('Anno',[2026,2027],index=0)
if year not in st.session_state.get('_ensured_years', set()):
    # ensure_jobs(year)
    st.session_state.setdefault('_ensured_years', set()).add(year)
    invalidate_data()
p=dataset(year); now=datetime.now()
if p.empty and area!='📊 Amministratore': st.warning('Nessuna manutenzione disponibile. Chiedi all’amministratore di importare il file Excel.'); st.stop()

if area=='📊 Amministratore':
    admin_password = st.secrets.get('ADMIN_PASSWORD')
    if not admin_password:
        st.error('Configura ADMIN_PASSWORD nei Secrets di Streamlit per usare l’area amministratore.')
        st.stop()
    entered = st.sidebar.text_input('Password amministratore', type='password', key='admin_password')
    if entered != admin_password:
        st.info('Inserisci la password amministratore per accedere.')
        st.stop()
    st.subheader('🗓️ Generazione controllata programmazione 2027')
    st.caption('Copia dal 2026 soltanto nel database di PROVA. S, N e ? restano invariati; nessuna modifica allo storico 2026. Se un lotto si interrompe, la procedura può essere ripetuta dopo i controlli.')
    confirm_2027 = st.checkbox('Confermo la generazione della programmazione 2027 SOLO nel database di PROVA', key='confirm_generate_2027')
    if st.button('GENERA PROGRAMMAZIONE 2027 DA 2026', disabled=not confirm_2027, key='generate_2027'):
        try:
            with st.spinner('Controllo e generazione 2027 in corso. Non chiudere la pagina.'):
                nplans,njobs,totaljobs=generate_2027_from_2026()
        except Exception as exc:
            st.error('Generazione bloccata o incompleta: ' + str(exc))
        else:
            st.success(f'Programmazione 2027 pronta: {nplans} piani e {njobs} manutenzioni aggiunti. Totale manutenzioni 2027 previste: {totaljobs}.')
            st.rerun()
    st.divider()
    st.subheader('📅 Programmazione S/N e archivio Excel')
    st.caption('S = manutenzione programmata · N = non programmata · ? = da verificare. Le modifiche sono riservate all’amministratore.')
    plan_rows = pd.DataFrame(all_rows('programmazione', '*'))
    plant_rows = pd.DataFrame(all_rows('impianti', 'id,codice,codice_cliente,cliente,matricola,numero_fabbrica,comune,indirizzo,provincia,tecnico_assegnato,tipo_manutenzione,progressivo,note'))
    month_cols = [name.lower() for name,_,_ in MESI]
    if not plant_rows.empty:
        plans_year = plan_rows[plan_rows['anno'].astype(int) == year].copy() if not plan_rows.empty else pd.DataFrame()
        if not plans_year.empty:
            plans_year = plans_year.sort_values('id').drop_duplicates('impianto_id',keep='last')
            archive = plant_rows.merge(plans_year[['impianto_id'] + month_cols],left_on='id',right_on='impianto_id',how='left')
        else:
            archive = plant_rows.copy()
            for col in month_cols: archive[col] = None
        export = archive[['codice','codice_cliente','cliente','matricola','numero_fabbrica','comune','indirizzo','provincia','tecnico_assegnato','tipo_manutenzione','progressivo','note']].copy()
        export.columns = ['Codice','Codice cliente','Cliente','Matricola','Numero fabbrica','Comune','Indirizzo','Provincia','Tecnico','Tipo manutenzione','Progressivo','Note']
        for name,_,_ in MESI:
            export[name.upper()] = archive[name.lower()].map(lambda x: '?' if pd.isna(x) else ('S' if x is True or str(x).lower()=='true' else 'N'))
        export['Visite anno'] = (export[[name.upper() for name,_,_ in MESI]] == 'S').sum(axis=1)
        export_buf = io.BytesIO()
        with pd.ExcelWriter(export_buf,engine='openpyxl') as writer:
            export.to_excel(writer,index=False,sheet_name=f'Programmazione {year}')
        st.download_button('⬇️ Scarica archivio Excel aggiornato',export_buf.getvalue(),file_name=f'TESYX_programmazione_{year}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key=f'export_plan_{year}')
        st.caption('L’Excel contiene i dati e i mesi S/N presenti su Supabase. Scaricarlo non modifica nulla.')
        q_plan=st.text_input('🔎 Cerca impianto da programmare (codice, matricola, cliente, indirizzo)')
        found_plan=archive.copy()
        if q_plan.strip():
            fields=found_plan[['codice','matricola','cliente','indirizzo','comune']].fillna('').astype(str)
            found_plan=found_plan[fields.apply(lambda col:col.str.contains(q_plan.strip(),case=False,regex=False)).any(axis=1)]
        st.caption(f'Impianti trovati: {len(found_plan)}')
        if not found_plan.empty:
            ids=found_plan['id'].astype(int).tolist()
            labels={int(r['id']):f"{r['codice']} · {r['indirizzo']} · {r['cliente']} · {r['tecnico_assegnato']}" for _,r in found_plan.iterrows()}
            selected_id=st.selectbox('Impianto da modificare',ids,format_func=lambda i:labels[i],index=None,placeholder='Seleziona impianto',key='edit_plan_plant')
            if selected_id is not None:
                row=found_plan[found_plan['id']==selected_id].iloc[0]
                with st.form(f'edit_plan_{year}_{selected_id}'):
                    cols=st.columns(6)
                    new_months={}
                    for j,(name,full_name,num) in enumerate(MESI):
                        old='?' if pd.isna(row[name.lower()]) else ('S' if row[name.lower()] is True or str(row[name.lower()]).lower()=='true' else 'N')
                        chosen=cols[j%6].selectbox(full_name.upper(),['S','N','?'],index=['S','N','?'].index(old),key=f'plan_{year}_{selected_id}_{name}')
                        new_months[name.lower()]=True if chosen=='S' else False if chosen=='N' else None
                    reason=st.text_input('Motivo della modifica (obbligatorio)')
                    confirm=st.checkbox('Confermo di voler aggiornare la programmazione online')
                    save=st.form_submit_button('SALVA PROGRAMMAZIONE SU SUPABASE',type='primary')
                if save:
                    if not reason.strip() or not confirm:
                        st.error('Inserisci il motivo e conferma prima di salvare.')
                    else:
                        try:
                            result=update_plan_safely(selected_id,year,new_months,reason.strip())
                        except Exception as exc:
                            st.error(f'Operazione non completata: {exc}')
                        else:
                            st.success(result)
                            st.rerun()
    st.divider()
    st.subheader('🧪 Anteprima nuovo archivio Excel (senza scrivere su Supabase)')
    st.caption('Legge il Foglio1 del file definitivo per gestionale.xlsx. Non importa, cancella o aggiorna alcun dato.')
    preview_upload = st.file_uploader('Seleziona definitivo per gestionale.xlsx', type=['xlsx'], key='new_archive_preview')
    if preview_upload is not None:
        try:
            raw = pd.read_excel(preview_upload, sheet_name='Foglio1', header=2, dtype=str).fillna('')
            raw.columns = [str(x).replace('\n',' ').strip().upper() for x in raw.columns]
            needed = ['COD. CLIENTE','RAG. SOCIALE','NUMERO MATRICOLA','NUMERO FABBRICA',"CITTA'",'UBICAZIONE','PROV.','TIPO MANUT.','MANUTENTORE']
            missing = [x for x in needed if x not in raw.columns]
            if missing:
                st.error('Intestazioni non riconosciute: ' + ', '.join(missing))
            else:
                raw = raw[(raw['RAG. SOCIALE'].str.strip() != '') | (raw['UBICAZIONE'].str.strip() != '')].copy()
                raw['TIPO MANUT.'] = raw['TIPO MANUT.'].str.strip().str.upper()
                raw['MESI RICONOSCIUTI'] = raw['TIPO MANUT.'].map(lambda x: months_for_cadence(x))
                raw['VERIFICA'] = raw.apply(lambda r: 'DA DEFINIRE' if not r['TIPO MANUT.'] else ('TIPO NON STANDARD' if r['MESI RICONOSCIUTI'] is None else 'AUTOMATICA'), axis=1)
                st.write(f"Impianti letti: **{len(raw)}** · Cadenze mancanti: **{int((raw['VERIFICA']=='DA DEFINIRE').sum())}** · Tipi non standard: **{int((raw['VERIFICA']=='TIPO NON STANDARD').sum())}**")
                st.dataframe(raw['TIPO MANUT.'].value_counts(dropna=False).rename_axis('Tipo').reset_index(name='Impianti'), hide_index=True)
                uncertain = raw[raw['VERIFICA']!='AUTOMATICA']
                st.markdown('**Impianti da programmare manualmente**')
                st.dataframe(uncertain[['COD. CLIENTE','RAG. SOCIALE','NUMERO MATRICOLA','UBICAZIONE','TIPO MANUT.','MANUTENTORE','VERIFICA']], hide_index=True, use_container_width=True)
                st.info('Anteprima completata: nessun dato è stato scritto su Supabase. Le cadenze manuali si potranno salvare nel gestionale dopo la verifica e la migrazione controllata.')
        except Exception as exc:
            st.error('Impossibile leggere il nuovo Excel: ' + str(exc))
    st.divider()
    st.subheader('💾 Backup di sicurezza prima della sostituzione')
    st.caption('Scarica una copia completa delle tabelle del gestionale prima di qualsiasi modifica. Il backup non scrive su Supabase.')
    if st.button('Prepara backup completo Supabase', key='backup_official_archive'):
        try:
            backup_tables=['impianti_v2','programmazione_v2','manutenzioni_v2',
                           'tecnici_v2','stati_impianti_v2','storico_annullamenti_v2',
                           'storico_programmazione_v2']
            buffer=io.BytesIO()
            counts={}; errors={}
            with zipfile.ZipFile(buffer,'w',zipfile.ZIP_DEFLATED) as archive:
                for table_name in backup_tables:
                    try:
                        records=fetch_rows(table_name,'*')
                        counts[table_name]=len(records)
                        archive.writestr(table_name+'.json',json.dumps(records,ensure_ascii=False,default=str,indent=2))
                    except Exception as exc:
                        errors[table_name]=str(exc)
                manifest={'created_at':datetime.now().isoformat(), 'counts':counts, 'errors':errors,
                          'warning':'Backup per consultazione; non ripristinare automaticamente senza verifica dei vincoli.'}
                archive.writestr('MANIFEST.json',json.dumps(manifest,ensure_ascii=False,indent=2))
            st.session_state['archive_backup_bytes']=buffer.getvalue()
            st.session_state['archive_backup_manifest']=manifest
        except Exception as exc:
            st.error('Backup non riuscito: '+str(exc))
    if st.session_state.get('archive_backup_bytes'):
        manifest=st.session_state['archive_backup_manifest']
        if manifest['errors']:
            st.error('Backup incompleto: '+', '.join(manifest['errors'])+'. Non procedere alla sostituzione.')
        else:
            st.success('Backup completo pronto. Scaricalo e conservalo sul computer.')
        st.dataframe(pd.DataFrame([{'Tabella':k,'Righe':v} for k,v in manifest['counts'].items()]),hide_index=True)
        st.download_button('⬇️ Scarica backup completo ZIP',data=st.session_state['archive_backup_bytes'],
                           file_name='TESYX_backup_prima_sostituzione.zip',mime='application/zip',
                           key='download_full_backup')
    st.divider()
    st.subheader('🧭 Preparazione sostituzione archivio ufficiale')
    st.caption('Confronto di sicurezza: il nuovo Excel prevale per assegnazioni e cadenze. Nessuna scrittura su Supabase in questa fase.')
    official_file=st.file_uploader('Carica definitivo per gestionale.xlsx per preparare la sostituzione',type=['xlsx'],key='official_replace_file')
    if official_file is not None:
        try:
            source=pd.read_excel(official_file,sheet_name='Foglio1',header=2,dtype=str).fillna('')
            source.columns=[str(c).replace('\n',' ').strip().upper() for c in source.columns]
            required=['COD. CLIENTE','RAG. SOCIALE','NUMERO MATRICOLA','NUMERO FABBRICA',"CITTA'",'UBICAZIONE','PROV.','TIPO MANUT.','MANUTENTORE']
            if any(c not in source.columns for c in required):
                raise ValueError('Colonne mancanti: '+', '.join(c for c in required if c not in source.columns))
            source=source[(source['RAG. SOCIALE'].str.strip()!='') | (source['UBICAZIONE'].str.strip()!='')].copy()
            source['TIPO MANUT.']=source['TIPO MANUT.'].str.strip().str.upper()
            source['MESI']=source['TIPO MANUT.'].map(months_for_cadence)
            existing=fetch_rows('impianti','id,codice,matricola,numero_fabbrica,indirizzo,comune,tecnico_assegnato,tipo_manutenzione')
            def norm(value):
                return re.sub(r'[^A-Z0-9]','',clean(value).upper())
            def keys(record):
                return (norm(record.get('matricola')),norm(record.get('numero_fabbrica')))
            by_mat={};by_fab={}
            for plant in existing:
                mat,fab=keys(plant)
                if mat:by_mat.setdefault(mat,[]).append(plant)
                if fab:by_fab.setdefault(fab,[]).append(plant)
            comparisons=[];seen={}
            for _,r in source.iterrows():
                mat=norm(r['NUMERO MATRICOLA']);fab=norm(r['NUMERO FABBRICA'])
                a=by_mat.get(mat,[]) if mat else []
                b=by_fab.get(fab,[]) if fab else []
                candidates={v['id']:v for v in a+b}
                both={v['id'] for v in a}&{v['id'] for v in b} if mat and fab else set()
                if len(both)==1: candidate=candidates[next(iter(both))];status='CORRISPONDENZA CERTA'
                elif len(candidates)==1 and (not a or not b):candidate=next(iter(candidates.values()));status='DA VERIFICARE - UN SOLO IDENTIFICATIVO'
                elif not candidates:candidate=None;status='NUOVO / NON TROVATO'
                else:candidate=None;status='AMBIGUO - NON ABBINARE'
                if candidate:seen[candidate['id']]=seen.get(candidate['id'],0)+1
                comparisons.append({'Codice cliente':clean(r['COD. CLIENTE']),'Cliente':clean(r['RAG. SOCIALE']),'Matricola':clean(r['NUMERO MATRICOLA']),'Numero fabbrica':clean(r['NUMERO FABBRICA']),'Comune':clean(r["CITTA'"]),'Indirizzo':clean(r['UBICAZIONE']),'Tecnico nuovo':clean(r['MANUTENTORE']),'Cadenza nuova':clean(r['TIPO MANUT.']),'Mesi previsti':','.join(map(str,r['MESI'])) if isinstance(r['MESI'],list) else '', 'ID vecchio proposto':candidate['id'] if candidate else '', 'Tecnico vecchio':candidate.get('tecnico_assegnato','') if candidate else '', 'Esito':status})
            for item in comparisons:
                iid=item['ID vecchio proposto']
                if iid and seen[iid]>1:item['Esito']='AMBIGUO - PIÙ RIGHE SULLO STESSO ID'
            result=pd.DataFrame(comparisons)
            st.write('Nuovo archivio:',len(result),'impianti · Attuale Supabase:',len(existing),'impianti')
            st.dataframe(result['Esito'].value_counts().rename_axis('Esito').reset_index(name='Impianti'),hide_index=True)
            st.dataframe(result[result['Esito']!='CORRISPONDENZA CERTA'].head(100),hide_index=True,use_container_width=True)
            st.download_button('⬇️ Scarica confronto completo CSV',data=result.to_csv(index=False,sep=';').encode('utf-8-sig'),file_name='confronto_archivio_ufficiale.csv',mime='text/csv')
            st.subheader('🚀 Attivazione archivio ufficiale (senza cancellare lo storico)')
            st.info('Il vecchio archivio e tutti i collegamenti storici rimangono nel database. Si inserisce un archivio ufficiale separato; soltanto alla fine viene attivato per tecnici e amministratore.')
            if official_active():
                st.success('ARCHIVIO UFFICIALE GIÀ ATTIVO. Non ripetere l’importazione.')
            else:
                st.warning('Operazione reale su Supabase: prima di procedere conserva il backup ZIP già scaricato.')
                confirm=st.checkbox('Confermo di aver conservato il backup ZIP e di voler attivare i 2.399 impianti ufficiali',key='confirm_real_migration')
                if st.button('IMPORTA E ATTIVA ARCHIVIO UFFICIALE',disabled=not confirm,key='real_migration'):
                    try:
                        with st.spinner('Importazione reale in corso. Non chiudere la pagina.'):
                            counts=activate_official_archive(source)
                        st.success('Archivio ufficiale attivato. Ricarica la pagina per visualizzarlo.')
                        st.json(counts)
                    except Exception as ex:
                        st.error('Importazione interrotta: '+str(ex))
                        st.warning('Non ripetere senza controllo: l’archivio precedente resta operativo finché l’attivazione non è completata.')
        except Exception as exc:
            st.error('Preparazione non completata: '+str(exc))
    st.divider()
    st.subheader('📥 Importazione archivio impianti')
    count=db().table('impianti_v2').select('id',count='exact').limit(1).execute().count or 0
    st.write(f'Impianti presenti nel nuovo archivio: **{count}**')
    if count==0:
        upload=st.file_uploader('Carica Archivio_unico_VERIFICATO_test_ripristinato.xlsx',type=['xlsx'])
        if upload is not None:
            try:
                preview=import_frame(upload)
                month_cols_excel=['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic']
                unknown=(preview[month_cols_excel]=='?').sum(axis=1)
                st.write(f"Impianti: **{len(preview)}** · Impianti con ?: **{int((unknown>0).sum())}** · Celle ?: **{int(unknown.sum())}**")
                st.warning('I mesi ? resteranno DA VERIFICARE e non genereranno manutenzioni. La procedura scrive nel database di PROVA e non può essere annullata automaticamente.')
                st.dataframe(preview.loc[unknown>0,['Cliente','Matricola','Indirizzo','Esito controllo']+month_cols_excel].head(30),hide_index=True)
                st.dataframe(preview[['Cliente','Matricola','Indirizzo','Tecnico proposto']].head(12),hide_index=True)
                agree=st.checkbox('Ho verificato il file e confermo l’importazione esclusivamente nel database di PROVA')
                if st.button('IMPORTA NEL DATABASE DI PROVA',type='primary',disabled=not agree):
                    with st.spinner('Importazione in corso. Non chiudere la pagina.'):
                        n,j,t=import_new_archive(preview)
                    st.success(f'Importati {n} impianti, {j} manutenzioni e {t} tecnici.')
                    st.rerun()
            except Exception as exc:
                st.error('Importazione non eseguita o incompleta: '+str(exc))
    else:
        st.info('Impianti già presenti: nessuna nuova importazione degli impianti.')
        st.markdown('#### Ripresa controllata dell’importazione interrotta (solo 2026)')
        recovery_upload=st.file_uploader('Seleziona lo stesso archivio Excel ripristinato',type=['xlsx'],key='recover_excel')
        if recovery_upload is not None:
            try:
                recovery_df=import_frame(recovery_upload)
                if len(recovery_df)!=2547:
                    st.error('Numero di righe inatteso: ripresa bloccata.')
                else:
                    st.caption('Il recupero confronta tutti gli impianti e i dati già presenti. Non sovrascrive programmazioni o manutenzioni.')
                    confirm_recovery=st.checkbox('Confermo la ripresa SOLO nel database di PROVA',key='confirm_recovery')
                    if st.button('COMPLETA IMPORTAZIONE DI PROVA',disabled=not confirm_recovery,key='recover_button'):
                        with st.spinner('Verifica e completamento in corso...'):
                            np,nj,nt=recover_partial_import(recovery_df)
                        st.success(f'Operazione completata: {np} programmazioni e {nj} manutenzioni mancanti aggiunte; {nt} tecnici verificati.')
                        st.rerun()
            except Exception as exc:
                st.error('Ripresa bloccata o incompleta: '+str(exc))
    st.divider()
    st.subheader('🚦 Stato impianti: fermo, sospeso e riattivazione')
    st.caption('Le NOTE originali non vengono modificate. La riattivazione ha effetto dal mese corrente, senza arretrati.')
    plants = pd.DataFrame(all_rows('impianti','id,codice,cliente,indirizzo,comune,progressivo,tecnico_assegnato,note'))
    if not plants.empty:
        status_map = get_status_overrides()
        plants['stato_attuale'] = plants.apply(lambda r:status_map.get(int(r['id']),{}).get('stato') or initial_status(r.get('note')),axis=1)
        suspended = plants[plants['stato_attuale']!='ATTIVO']
        st.write(f'Impianti non attivi: **{len(suspended)}**')
        st.dataframe(suspended[['id','progressivo','cliente','indirizzo','tecnico_assegnato','stato_attuale','note']],hide_index=True,use_container_width=True)
        query_plant=st.text_input('Cerca impianto da modificare (indirizzo, cliente, progressivo)')
        found=plants.copy()
        if query_plant.strip():
            found=found[found[['cliente','indirizzo','comune','progressivo']].fillna('').astype(str).apply(lambda x:x.str.contains(query_plant,case=False,regex=False)).any(axis=1)]
        if not found.empty:
            ids=found['id'].astype(int).tolist()
            labels={int(r['id']):f"PROGR {r['progressivo'] or '-'} · {r['indirizzo']} · {r['cliente']} · {r['tecnico_assegnato']}" for _,r in found.iterrows()}
            selected_plant=st.selectbox('Seleziona impianto',ids,format_func=lambda i:labels[i],index=None,placeholder='Scegli un impianto')
            if selected_plant is not None:
                selected_row=plants[plants.id==selected_plant].iloc[0]
                st.info(f"Stato attuale: {selected_row['stato_attuale']} · Nota originale: {clean(selected_row.get('note')) or 'Nessuna'}")
                options=['ATTIVO','FERMO','SOSPESO','DISDETTO','CHIUSO']
                new_status=st.selectbox('Nuovo stato',options,index=options.index(selected_row['stato_attuale']),key=f'status_{selected_plant}')
                reason=st.text_input('Motivo o nota amministrativa',key=f'reason_{selected_plant}')
                if st.button('SALVA STATO IMPIANTO',disabled=new_status==selected_row['stato_attuale']):
                    save_plant_status(selected_plant,new_status,reason)
                    st.success('Stato aggiornato.'); st.rerun()
    st.divider()
    if p.empty:
        st.info('Carica prima l’Excel per visualizzare la dashboard.')
        st.stop()
    st.subheader('📊 Dashboard Amministratore')
    month_num=st.selectbox('Periodo',range(1,13),index=now.month-1,format_func=lambda x:MESI[x-1][1])
    pm=p[(p.mese_competenza==month_num)&(p.stato_ui.isin(['DA FARE','SCADUTA','ESEGUITA']))].copy(); overdue=p[(p.stato_ui=='SCADUTA') & (p.mese_competenza<month_num if year==now.year else True)]
    total=len(pm); done=(pm.stato_ui=='ESEGUITA').sum(); todo=(pm.stato_ui=='DA FARE').sum(); sem=int(pm.semestrale.fillna(False).astype(bool).sum())
    a,b,c,d,e=st.columns(5); a.metric('Programmate',total); b.metric('🟢 Eseguite',done); c.metric('🟠 Da fare',todo); d.metric('🔴 Scadute precedenti',len(overdue)); e.metric('Semestrali eseguite',sem)
    if total: st.progress(done/total,text=f'Avanzamento {done/total*100:.1f}%')
    st.markdown('### 📅 Scadenze biennali')
    # Scadenze da tutti gli anni, non soltanto dall'anno selezionato.
    biennali_jobs = pd.DataFrame(all_rows('manutenzioni','impianto_id,stato,biennale,scadenza_biennale'))
    biennali = (biennali_jobs[(biennali_jobs['stato']=='ESEGUITA') & (biennali_jobs['biennale'].fillna(False).astype(bool)) & biennali_jobs['scadenza_biennale'].notna()].copy() if not biennali_jobs.empty else pd.DataFrame())
    if not biennali.empty:
        imps_b = pd.DataFrame(all_rows('impianti','id,progressivo,cliente,indirizzo,tecnico_assegnato'))
        biennali = biennali.merge(imps_b,left_on='impianto_id',right_on='id',how='left').rename(columns={'tecnico_assegnato':'assegnato'})
    if biennali.empty:
        st.info('Nessuna manutenzione biennale registrata.')
    else:
        # Una sola scadenza per impianto: la biennale eseguita più di recente.
        biennali['scad_dt'] = pd.to_datetime(biennali['scadenza_biennale'], errors='coerce')
        biennali = biennali.sort_values('scad_dt').drop_duplicates('impianto_id', keep='last')
        biennali['situazione'] = biennali['scad_dt'].apply(lambda d: 'SCADUTA' if pd.notna(d) and d.date() < date.today() else ('ENTRO 60 GIORNI' if pd.notna(d) and d.date() <= date.today() + relativedelta(days=60) else 'PROGRAMMATA'))
        st.dataframe(biennali[['progressivo','cliente','indirizzo','assegnato','scadenza_biennale','situazione']].rename(columns={'progressivo':'Progressivo','cliente':'Cliente','indirizzo':'Indirizzo','assegnato':'Tecnico','scadenza_biennale':'Scadenza biennale','situazione':'Situazione'}), hide_index=True,use_container_width=True)
    st.markdown('### Andamento per tecnico assegnatario')
    rows=[]
    for t,g in pm.groupby('assegnato',dropna=False):
        n=len(g); ex=(g.stato_ui=='ESEGUITA').sum(); rows.append([t,n,ex,n-ex,round(ex/n*100,1) if n else 0])
    st.dataframe(pd.DataFrame(rows,columns=['Tecnico','Previste','Eseguite','Mancanti','Avanzamento %']),hide_index=True,use_container_width=True)
    st.markdown('### Lavoro realmente eseguito dai tecnici')
    exm=pm[pm.stato_ui=='ESEGUITA'].copy()
    if exm.empty: st.info('Nessuna manutenzione registrata nel periodo.')
    else:
        st.dataframe(exm.groupby('tecnico_esecutore').size().reset_index(name='Manutenzioni eseguite').sort_values('Manutenzioni eseguite',ascending=False),hide_index=True,use_container_width=True)
        exm['giorno']=pd.to_datetime(exm.eseguita_il,errors='coerce').dt.date
        st.markdown('### Attività giornaliera')
        st.dataframe(exm.groupby(['giorno','tecnico_esecutore']).size().reset_index(name='Manutenzioni').sort_values(['giorno','tecnico_esecutore'],ascending=[False,True]),hide_index=True,use_container_width=True)
    st.markdown('### 📋 Dettaglio manutenzioni eseguite e note')
    completed = p[p['stato']=='ESEGUITA'].copy()
    search_done = st.text_input('🔎 Cerca nelle manutenzioni eseguite (codice, indirizzo, cliente, note)')
    if search_done:
        fields = completed[['codice','indirizzo','cliente','note']].fillna('').astype(str)
        completed = completed[fields.apply(lambda col: col.str.contains(search_done,case=False,regex=False)).any(axis=1)]
    if completed.empty:
        st.info('Nessuna manutenzione eseguita corrispondente alla ricerca.')
    else:
        completed = completed.sort_values('eseguita_il',ascending=False,na_position='last')
        display = completed[['codice','indirizzo','cliente','mese_nome','anno_competenza','tecnico_esecutore','eseguita_il','semestrale','biennale','scadenza_biennale','note']].copy()
        display.columns=['Codice','Indirizzo','Cliente','Mese','Anno','Tecnico esecutore','Data esecuzione','Semestrale','Biennale','Scadenza biennale','Note']
        st.dataframe(display.fillna(''),hide_index=True,use_container_width=True)
        st.markdown('#### ↩ Annullamento di una manutenzione errata')
        choices = {int(r['id']): f"{r['codice']} · {r['indirizzo']} · {r['mese_nome']} {r['anno_competenza']} · {r['eseguita_il']}" for _,r in completed.iterrows()}
        chosen = st.selectbox('Seleziona la registrazione da annullare',options=list(choices),format_func=lambda k: choices[k],index=None,placeholder='Seleziona una manutenzione')
        if chosen is not None:
            st.warning('L’intervento tornerà da eseguire (o scaduto, secondo il mese). I dati della registrazione annullata saranno conservati nello storico.')
            reason = st.text_input('Motivo dell’annullamento (obbligatorio)')
            confirmed = st.checkbox('Confermo di voler annullare questa registrazione',key=f'confirm_cancel_{chosen}')
            if st.button('↩ ANNULLA MANUTENZIONE',disabled=not (confirmed and reason.strip()),type='secondary'):
                try:
                    cancel_job(chosen,reason.strip())
                except Exception as exc:
                    st.error('Annullamento non effettuato. Verifica che la tabella storico_annullamenti sia stata creata in Supabase.')
                    st.exception(exc)
                else:
                    st.success('Manutenzione annullata e storico salvato.')
                    st.rerun()
    st.markdown('### 🗂️ Storico annullamenti')
    try:
        audit = pd.DataFrame(all_rows('storico_annullamenti','*'))
        if audit.empty:
            st.caption('Nessun annullamento registrato.')
        else:
            st.dataframe(audit.sort_values('annullato_il',ascending=False),hide_index=True,use_container_width=True)
    except Exception:
        st.info('Lo storico sarà disponibile dopo aver creato la tabella storico_annullamenti in Supabase.')
    st.markdown('### 🔴 Criticità')
    if overdue.empty: st.success('Nessuna manutenzione arretrata.')
    else: st.dataframe(overdue[['mese_nome','codice','cliente','indirizzo','comune','assegnato']],hide_index=True,use_container_width=True)
else:
    if year == 2026:
        st.info('2026 operativo: Cristian Malfatti da ottobre; gli altri tecnici da novembre. I mesi precedenti restano NON AVVIATA.')
    techs=['TUTTI']+sorted(p.assegnato.dropna().unique().tolist())
    cristian_idx = 0
    tech=st.sidebar.selectbox('Tecnico assegnatario',techs,index=cristian_idx)
    stato=st.sidebar.selectbox('Stato',['TUTTI','SCADUTA','DA FARE','ESEGUITA','DA NON FARE','NON AVVIATA']); mese=st.sidebar.selectbox('Mese',['TUTTI']+[x[1] for x in MESI],index=(now.month if now.month<=12 else 0))
    f=p.copy()
    if tech!='TUTTI': f=f[f.assegnato==tech]
    if stato!='TUTTI': f=f[f.stato_ui==stato]
    if mese!='TUTTI': f=f[f.mese_nome==mese]
    c1,c2,c3,c4=st.columns(4); c1.metric('Programmate',int(f.stato_ui.isin(['DA FARE','SCADUTA','ESEGUITA']).sum())); c2.metric('🟢 Eseguite',(f.stato_ui=='ESEGUITA').sum()); c3.metric('🟠 Da fare',(f.stato_ui=='DA FARE').sum()); c4.metric('🔴 Scadute',(f.stato_ui=='SCADUTA').sum())
    if f.stato_ui.isin(['DA FARE','SCADUTA','ESEGUITA']).sum():
        active_count=int(f.stato_ui.isin(['DA FARE','SCADUTA','ESEGUITA']).sum())
        st.progress(float((f.stato_ui=='ESEGUITA').sum()/active_count),text=f"Avanzamento {((f.stato_ui=='ESEGUITA').sum()/active_count)*100:.1f}%")
    st.subheader('Manutenzioni')
    st.caption('🔴 Le manutenzioni scadute restano evidenziate finché non vengono registrate.')
    q=st.text_input('🔎 Cerca progressivo, codice, matricola, cliente, indirizzo o comune')
    if q: f=f[f[['progressivo','codice','matricola','cliente','indirizzo','comune']].fillna('').astype(str).apply(lambda x:x.str.contains(q,case=False,regex=False)).any(axis=1)]
    if date.today() < date(year, 1, 1):
        st.info(f'Le manutenzioni {year} sono consultabili, ma la registrazione sarà disponibile dal 1° gennaio {year}.')
        st.session_state.pop('selected_job', None)
    if date.today() >= date(year, 1, 1) and 'selected_job' in st.session_state:
        rr=p[p['id']==st.session_state.selected_job]
        if not rr.empty and int(rr.iloc[0]['anno_competenza']) == year and rr.iloc[0]['stato_ui'] in ('DA FARE','SCADUTA'):
            r=rr.iloc[0]
            with st.container(border=True):
                st.markdown(f"### Manutenzione — {r['indirizzo']}"); st.write(f"**{r['cliente']}** · {r['comune']}"); st.caption(f"Competenza: {r['mese_nome']} {r['anno_competenza']} · Assegnato a: {r['assegnato']} · Cod. {r['codice']}")
                cc1,cc2=st.columns(2); ops=operators(); default=ops.index(r['assegnato']) if r['assegnato'] in ops else 0; oper=cc1.selectbox('Tecnico che ha eseguito',ops,index=default); data_exec=cc2.date_input('Data esecuzione',value=date.today())
                sem=st.checkbox('Semestrale eseguita'); biennale=st.checkbox('Biennale eseguita'); nota=st.text_area('Note (facoltative)')
                if biennale: st.info(f'Prossima scadenza biennale: {(data_exec + relativedelta(months=24)).strftime("%d/%m/%Y")}')
                b1,b2=st.columns([3,1])
                if b1.button('CONFERMA MANUTENZIONE',type='primary',use_container_width=True):
                    try:
                        close_job(r['id'],oper,data_exec,sem,biennale,nota)
                    except Exception as exc: st.error(str(exc))
                    else:
                        del st.session_state.selected_job; st.success('Manutenzione registrata.'); st.rerun()
                if b2.button('Annulla',use_container_width=True): del st.session_state.selected_job; st.rerun()
    if f.empty: st.info('Nessuna manutenzione con i filtri selezionati.')
    else:
        order={'SCADUTA':0,'DA FARE':1,'DA NON FARE':2,'NON AVVIATA':3,'ESEGUITA':4}
        mode=st.radio('Ordina il giro', ['Per via e civico', 'Per vicinanza geografica', 'Ordine tradizionale'], horizontal=True, key='route_mode')
        f=f.assign(ord=f.stato_ui.map(order))
        if mode in ('Per via e civico', 'Per vicinanza geografica'):
            parts=f['indirizzo'].apply(street_sort_parts)
            f=f.assign(via_giro=parts.map(lambda x:x[0]),civico_giro=parts.map(lambda x:x[1]),lettera_giro=parts.map(lambda x:x[2]))
            f=f.sort_values(['assegnato','mese_competenza','comune','via_giro','civico_giro','lettera_giro','ord','progressivo'],kind='stable',na_position='last')
        else:
            f=f.sort_values(['ord','mese_competenza','comune','indirizzo'],kind='stable')
        if mode == 'Per vicinanza geografica':
            st.caption('Calcolo sperimentale sulla distanza in linea d’aria, NON sui tempi di percorrenza stradali. Si consiglia di filtrare un solo tecnico e un solo mese.')
            route_key=(year, tech, mese, stato, q)
            if st.session_state.get('geo_route_key') != route_key:
                st.session_state.pop('geo_route_coords', None)
            if st.button('📍 Calcola giro geografico', key='calculate_geo_route'):
                if tech == 'TUTTI' or mese == 'TUTTI':
                    st.warning('Seleziona un solo tecnico e un solo mese prima di calcolare il giro.')
                else:
                    addresses=list(dict.fromkeys((clean(r['comune']).upper(), clean(r['indirizzo']).upper(), clean(r.get('provincia')).upper()) for _,r in f.iterrows()))
                    if len(addresses)>80:
                        st.warning('Troppi indirizzi per una singola ricerca. Filtra il giro a massimo 80 indirizzi.')
                    else:
                        coords={}; errors=0
                        bar=st.progress(0, text='Ricerca posizioni degli impianti...')
                        for idx,key in enumerate(addresses):
                            try:
                                coords[key]=geocode_address(key[1],key[0],key[2])
                            except Exception:
                                coords[key]=None; errors+=1
                            bar.progress((idx+1)/len(addresses), text=f'Posizioni verificate: {idx+1}/{len(addresses)}')
                            if idx+1<len(addresses): time.sleep(1.1)  # rispetta il limite pubblico di Nominatim
                        st.session_state['geo_route_coords']=coords
                        st.session_state['geo_route_key']=route_key
                        if errors: st.warning(f'{errors} indirizzi non sono stati verificati per problemi di connessione.')
            if st.session_state.get('geo_route_key') == route_key and st.session_state.get('geo_route_coords'):
                f, missing=geographic_order(f,st.session_state['geo_route_coords'])
                st.info(f'Giro geografico calcolato. Indirizzi senza posizione certa: {missing}. Questi sono collocati alla fine del giro.')
            else:
                st.info('Premi «Calcola giro geografico» per ottenere la sequenza per vicinanza.')
        page_size=25; pages=max(1,(len(f)+page_size-1)//page_size); page=st.number_input('Pagina',1,pages,1,1) if pages>1 else 1; start=(int(page)-1)*page_size
        for _,r in f.iloc[start:start+page_size].iterrows():
            with st.container(border=True):
                a,b,c=st.columns([6,2,2]); icon='🔴' if r.stato_ui in ('SCADUTA','DA NON FARE') else ('🟢' if r.stato_ui=='ESEGUITA' else '🟠'); a.markdown(f"**{icon} {r['indirizzo']} — {r['comune']}**"); a.caption(f"{r['cliente']} · PROGR. {r['progressivo'] if pd.notna(r['progressivo']) else '-'} · Matr. {r['matricola']} · Assegnato: {r['assegnato']}"); b.markdown(f"**{r['mese_nome']} {r['anno_competenza']}**"); b.write(r.stato_ui)
                if r.stato_ui in ('DA FARE','SCADUTA') and date.today() < date(year, 1, 1):
                    c.caption(f'Registrazione disponibile dal 01/01/{year}')
                elif r.stato_ui in ('DA FARE','SCADUTA'):
                    with c.container(key='registration_button_'+str(r['id'])):
                        if st.button('REGISTRA MANUTENZIONE',key='do_'+str(r['id']),use_container_width=True,type='primary'): st.session_state.selected_job=int(r['id']); st.rerun()
                elif r.stato_ui=='DA NON FARE':
                    c.button('DA NON FARE',key='blocked_'+str(r['id']),disabled=True,use_container_width=True,type='primary')
                    c.caption(f"{r['stato_impianto']} · {clean(r.get('note_impianto'))[:120]}")
                elif r.stato_ui=='NON AVVIATA':
                    c.caption('Prima della data di avvio')
                else:
                    c.success(str(r.tecnico_esecutore or 'Eseguita')); c.caption(str(r.eseguita_il or ''))
                    if bool(r.semestrale): c.caption('☑ Semestrale')
                    if bool(r.biennale): c.caption(f'☑ Biennale · Prossima: {r.scadenza_biennale}')

st.caption('TESYX Manutenzioni v2 · Archivio separato · Dati condivisi online su Supabase.')
