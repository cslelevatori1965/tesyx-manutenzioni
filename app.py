import streamlit as st
import pandas as pd
from datetime import datetime, date
import hashlib
import re
from supabase import create_client

st.set_page_config(page_title="TESYX Manutenzioni", page_icon="🛠️", layout="wide")
st.markdown("""
<style>
@media (max-width: 700px) {
  .block-container {padding-top: 1rem; padding-left: .65rem; padding-right: .65rem;}
  div[data-testid="stMetric"] {border: 1px solid rgba(128,128,128,.25); padding: .45rem; border-radius: .65rem;}
  .stButton > button {min-height: 3rem; font-weight: 700;}
}
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
    return create_client(url, key)

def clean(v):
    if pd.isna(v): return ''
    return str(v).strip()


def months_for_cadence(code):
    code=clean(code).upper().replace(' ', '')
    if code in ('M','M1','MENSILE'): return list(range(1,13))
    match=re.fullmatch(r'([BTQS])([1-9][0-9]*)',code)
    if not match: return None
    period={'B':2,'T':3,'Q':4,'S':6}[match.group(1)]
    offset=int(match.group(2))
    if offset<1 or offset>period: return None
    return list(range(offset,13,period))


def import_frame(upload):
    df=pd.read_excel(upload,sheet_name=0,dtype=str).fillna('')
    df.columns=[str(c).strip().upper() for c in df.columns]
    required=['COD. CLIENTE','RAG. SOCIALE','NUMERO MATRICOLA',"CITTA'",'UBICAZIONE','TIPO MANUT.','MANUTENTORE']
    missing=[c for c in required if c not in df.columns]
    if missing: raise ValueError('Colonne mancanti: '+', '.join(missing))
    df=df[df[['RAG. SOCIALE','NUMERO MATRICOLA','UBICAZIONE']].ne('').any(axis=1)].copy()
    df=df[~df['RAG. SOCIALE'].str.contains(r'ASCENSORI\s+SERVIZI',case=False,na=False,regex=True)].copy()
    df['TIPO MANUT.']=df['TIPO MANUT.'].str.strip().str.upper()
    df['mesi']=df['TIPO MANUT.'].map(months_for_cadence)
    # Identificativo stabile per questa importazione, non derivato da PROGR. (riparte per tecnico).
    df['codice_import']=df.apply(lambda r:'ROB-'+hashlib.sha256('|'.join([clean(r.get('COD. CLIENTE')),clean(r.get('NUMERO MATRICOLA')),clean(r.get('UBICAZIONE')),clean(r.get("CITTA'"),),clean(r.get('NUMERO FABBRICA'))]).upper().encode()).hexdigest()[:20],axis=1)
    # Alcuni impianti distinti condividono indirizzo e hanno matricola assente.
    # Conserviamo ogni riga assegnando un suffisso stabile nell'ordine del file.
    ordinal=df.groupby('codice_import',sort=False).cumcount()
    repeated=df['codice_import'].duplicated(keep=False)
    df.loc[repeated,'codice_import']=df.loc[repeated,'codice_import']+'-'+(ordinal[repeated]+1).astype(str)
    return df


def import_new_archive(df):
    client=db()
    # Blocca importazioni ripetute; nessuna cancellazione automatica.
    count=client.table('impianti_v2').select('id',count='exact').limit(1).execute().count or 0
    if count: raise ValueError('Archivio già popolato. Importazione bloccata per proteggere i dati.')
    records=[]
    for _,r in df.iterrows():
        prog=clean(r.get('PROGR.'))
        try: prog=int(float(prog)) if prog else None
        except ValueError: prog=None
        records.append({'codice':r['codice_import'],'codice_cliente':clean(r.get('COD. CLIENTE')) or None,'cliente':clean(r.get('RAG. SOCIALE')) or None,'matricola':clean(r.get('NUMERO MATRICOLA')) or None,'numero_fabbrica':clean(r.get('NUMERO FABBRICA')) or None,'cap':clean(r.get('CAP')) or None,'comune':clean(r.get("CITTA'")) or None,'indirizzo':clean(r.get('UBICAZIONE')) or None,'provincia':clean(r.get('PROV.')) or None,'tipo_manutenzione':clean(r.get('TIPO MANUT.')) or None,'progressivo':prog,'tecnico_assegnato':clean(r.get('MANUTENTORE')) or 'Non assegnato','note':clean(r.get('NOTE')) or None})
    for i in range(0,len(records),150): client.table('impianti_v2').insert(records[i:i+150]).execute()
    id_map={r['codice']:r['id'] for r in all_rows('impianti','id,codice')}
    year=date.today().year
    plans=[]; jobs=[]
    for (_,r),rec in zip(df.iterrows(),records):
        months=r['mesi'] or []
        iid=id_map[rec['codice']]
        plan={'impianto_id':iid,'anno':year}
        for _,name,num in MESI: plan[name.lower()]=num in months
        plans.append(plan)
        for month in months:
            jobs.append({'impianto_id':iid,'anno_competenza':year,'mese_competenza':month,'tecnico_assegnato':rec['tecnico_assegnato'],'stato':'DA_FARE'})
    for i in range(0,len(plans),150): client.table('programmazione_v2').insert(plans[i:i+150]).execute()
    for i in range(0,len(jobs),150): client.table('manutenzioni_v2').insert(jobs[i:i+150]).execute()
    names=sorted(set(r['tecnico_assegnato'] for r in records if r['tecnico_assegnato']!='Non assegnato'))
    for name in names: client.table('tecnici_v2').upsert({'nome':name,'attivo':True},on_conflict='nome').execute()
    dataset.clear()
    return len(records),len(jobs),len(names)

def all_rows(table, cols='*'):
    s=db(); rows=[]; start=0; step=1000
    while True:
        r=s.table(globals()["table"](table)).select(cols).range(start,start+step-1).execute().data
        rows.extend(r)
        if len(r)<step: break
        start+=step
    return rows

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
                for _,nome,_ in MESI: x[nome.lower()]=bool(r[nome.lower()])
                vals.append(x)
            for i in range(0,len(vals),200): s.table('programmazione_v2').insert(vals[i:i+200]).execute()
            prs=pd.DataFrame(all_rows('programmazione','*'))
    existing={(x['impianto_id'],x['anno_competenza'],x['mese_competenza']) for x in all_rows('manutenzioni','impianto_id,anno_competenza,mese_competenza')}
    vals=[]
    if prs.empty: return
    for _,r in prs[prs.anno==year].iterrows():
        tech=imps.loc[imps.id==r.impianto_id,'tecnico_assegnato']; tech=tech.iloc[0] if len(tech) else None
        for _,nome,num in MESI:
            if r[nome.lower()] is True and (int(r.impianto_id),year,num) not in existing:
                vals.append({'impianto_id':int(r.impianto_id),'anno_competenza':year,'mese_competenza':num,'tecnico_assegnato':tech,'stato':'DA_FARE'})
    for i in range(0,len(vals),200): s.table('manutenzioni_v2').insert(vals[i:i+200]).execute()

@st.cache_data(ttl=15)
def dataset(year):
    jobs=pd.DataFrame(all_rows('manutenzioni','id,impianto_id,anno_competenza,mese_competenza,tecnico_assegnato,tecnico_esecutore,stato,eseguita_il,semestrale,note'))
    imps=pd.DataFrame(all_rows('impianti','id,codice,matricola,cliente,indirizzo,comune,provincia,tecnico_assegnato'))
    if jobs.empty or imps.empty: return pd.DataFrame()
    p=jobs[jobs.anno_competenza==year].merge(imps,left_on='impianto_id',right_on='id',suffixes=('','_imp'))
    p['mese_nome']=p.mese_competenza.map({n:nome for _,nome,n in MESI})
    now=datetime.now(); current=now.month if year==now.year else (13 if year<now.year else 0)
    p['stato_ui']=p.apply(lambda r:'ESEGUITA' if r.stato=='ESEGUITA' else ('SCADUTA' if r.mese_competenza<current else 'DA FARE'),axis=1)
    p['assegnato']=p['tecnico_assegnato_imp'].fillna(p['tecnico_assegnato']).fillna('Non assegnato')
    return p

def operators():
    x=all_rows('tecnici','nome,attivo'); return sorted([r['nome'] for r in x if r.get('attivo',True)])

def close_job(job_id, oper, dt, sem, note):
    ts=datetime.combine(dt,datetime.now().time()).astimezone().isoformat()
    db().table('manutenzioni_v2').update({'stato':'ESEGUITA','tecnico_esecutore':oper,'eseguita_il':ts,'semestrale':bool(sem),'note':note}).eq('id',int(job_id)).execute()
    dataset.clear()

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
        'eseguita_il':None, 'semestrale':False, 'note':None
    }).eq('id',int(job_id)).eq('stato','ESEGUITA').execute()
    dataset.clear()

try:
    seed_if_needed()
except Exception as e:
    st.error('Impossibile inizializzare il database. Controlla i Secrets di Streamlit.'); st.exception(e); st.stop()

st.title('🛠️ TESYX · Gestione Manutenzioni')
area=st.sidebar.radio('Area',['👷 Tecnico','📊 Amministratore'], index=0)
year=st.sidebar.selectbox('Anno',[2026,2027],index=0)
ensure_jobs(year); p=dataset(year); now=datetime.now()
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
    st.subheader('📥 Importazione archivio impianti')
    count=db().table('impianti_v2').select('id',count='exact').limit(1).execute().count or 0
    st.write(f'Impianti presenti nel nuovo archivio: **{count}**')
    if count==0:
        upload=st.file_uploader('Carica Lista impianti Roberto – progressivi corretti (.xlsx)',type=['xlsx'])
        if upload is not None:
            try:
                preview=import_frame(upload)
                anomalies=preview[preview['mesi'].map(lambda x: x is None)]
                st.write(f'Righe valide: **{len(preview)}** · Cadenze da verificare: **{len(anomalies)}**')
                if len(anomalies):
                    st.warning('Le righe con cadenza non riconosciuta verranno importate senza manutenzioni programmate.')
                    st.dataframe(anomalies[['RAG. SOCIALE','NUMERO MATRICOLA','TIPO MANUT.']].head(50),hide_index=True)
                st.dataframe(preview[['RAG. SOCIALE','NUMERO MATRICOLA','TIPO MANUT.','MANUTENTORE']].head(12),hide_index=True)
                agree=st.checkbox('Confermo importazione iniziale nel nuovo archivio')
                if st.button('IMPORTA IMPIANTI',type='primary',disabled=not agree):
                    with st.spinner('Importazione in corso. Non chiudere la pagina.'):
                        n,j,t=import_new_archive(preview)
                    st.success(f'Importati {n} impianti, {j} manutenzioni e {t} tecnici.')
                    st.rerun()
            except Exception as exc:
                st.error('Importazione non eseguita o incompleta: '+str(exc))
    else:
        st.info('Archivio già inizializzato. Le importazioni successive richiederanno una funzione di aggiornamento protetto.')
    st.divider()
    if p.empty:
        st.info('Carica prima l’Excel per visualizzare la dashboard.')
        st.stop()
    st.subheader('📊 Dashboard Amministratore')
    month_num=st.selectbox('Periodo',range(1,13),index=now.month-1,format_func=lambda x:MESI[x-1][1])
    pm=p[p.mese_competenza==month_num].copy(); overdue=p[(p.stato_ui=='SCADUTA') & (p.mese_competenza<month_num if year==now.year else True)]
    total=len(pm); done=(pm.stato_ui=='ESEGUITA').sum(); todo=(pm.stato_ui=='DA FARE').sum(); sem=int(pm.semestrale.fillna(False).astype(bool).sum())
    a,b,c,d,e=st.columns(5); a.metric('Programmate',total); b.metric('🟢 Eseguite',done); c.metric('🟠 Da fare',todo); d.metric('🔴 Scadute precedenti',len(overdue)); e.metric('Semestrali eseguite',sem)
    if total: st.progress(done/total,text=f'Avanzamento {done/total*100:.1f}%')
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
        display = completed[['codice','indirizzo','cliente','mese_nome','anno_competenza','tecnico_esecutore','eseguita_il','semestrale','note']].copy()
        display.columns=['Codice','Indirizzo','Cliente','Mese','Anno','Tecnico esecutore','Data esecuzione','Semestrale','Note']
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
    techs=['TUTTI']+sorted(p.assegnato.dropna().unique().tolist())
    cristian_idx = 0
    tech=st.sidebar.selectbox('Tecnico assegnatario',techs,index=cristian_idx)
    stato=st.sidebar.selectbox('Stato',['TUTTI','SCADUTA','DA FARE','ESEGUITA']); mese=st.sidebar.selectbox('Mese',['TUTTI']+[x[1] for x in MESI],index=(now.month if now.month<=12 else 0))
    f=p.copy()
    if tech!='TUTTI': f=f[f.assegnato==tech]
    if stato!='TUTTI': f=f[f.stato_ui==stato]
    if mese!='TUTTI': f=f[f.mese_nome==mese]
    c1,c2,c3,c4=st.columns(4); c1.metric('Programmate',len(f)); c2.metric('🟢 Eseguite',(f.stato_ui=='ESEGUITA').sum()); c3.metric('🟠 Da fare',(f.stato_ui=='DA FARE').sum()); c4.metric('🔴 Scadute',(f.stato_ui=='SCADUTA').sum())
    if len(f): st.progress(float((f.stato_ui=='ESEGUITA').sum()/len(f)),text=f"Avanzamento {((f.stato_ui=='ESEGUITA').sum()/len(f))*100:.1f}%")
    st.subheader('Manutenzioni')
    st.caption('🔴 Le manutenzioni scadute restano evidenziate finché non vengono registrate.')
    q=st.text_input('🔎 Cerca codice, matricola, cliente, indirizzo o comune')
    if q: f=f[f[['codice','matricola','cliente','indirizzo','comune']].fillna('').astype(str).apply(lambda x:x.str.contains(q,case=False,regex=False)).any(axis=1)]
    if 'selected_job' in st.session_state:
        rr=p[p['id']==st.session_state.selected_job]
        if not rr.empty:
            r=rr.iloc[0]
            with st.container(border=True):
                st.markdown(f"### Manutenzione — {r['indirizzo']}"); st.write(f"**{r['cliente']}** · {r['comune']}"); st.caption(f"Competenza: {r['mese_nome']} {r['anno_competenza']} · Assegnato a: {r['assegnato']} · Cod. {r['codice']}")
                cc1,cc2=st.columns(2); ops=operators(); default=ops.index(r['assegnato']) if r['assegnato'] in ops else 0; oper=cc1.selectbox('Tecnico che ha eseguito',ops,index=default); data_exec=cc2.date_input('Data esecuzione',value=date.today())
                sem=st.checkbox('Semestrale eseguita'); nota=st.text_area('Note (facoltative)')
                b1,b2=st.columns([3,1])
                if b1.button('CONFERMA MANUTENZIONE',type='primary',use_container_width=True): close_job(r['id'],oper,data_exec,sem,nota); del st.session_state.selected_job; st.success('Manutenzione registrata.'); st.rerun()
                if b2.button('Annulla',use_container_width=True): del st.session_state.selected_job; st.rerun()
    if f.empty: st.info('Nessuna manutenzione con i filtri selezionati.')
    else:
        order={'SCADUTA':0,'DA FARE':1,'ESEGUITA':2}; f=f.assign(ord=f.stato_ui.map(order)).sort_values(['ord','mese_competenza','comune','indirizzo'])
        page_size=25; pages=max(1,(len(f)+page_size-1)//page_size); page=st.number_input('Pagina',1,pages,1,1) if pages>1 else 1; start=(int(page)-1)*page_size
        for _,r in f.iloc[start:start+page_size].iterrows():
            with st.container(border=True):
                a,b,c=st.columns([6,2,2]); icon='🔴' if r.stato_ui=='SCADUTA' else ('🟢' if r.stato_ui=='ESEGUITA' else '🟠'); a.markdown(f"**{icon} {r['indirizzo']} — {r['comune']}**"); a.caption(f"{r['cliente']} · Cod. {r['codice']} · Matr. {r['matricola']} · Assegnato: {r['assegnato']}"); b.markdown(f"**{r['mese_nome']} {r['anno_competenza']}**"); b.write(r.stato_ui)
                if r.stato_ui!='ESEGUITA':
                    if c.button('REGISTRA MANUTENZIONE',key='do_'+str(r['id']),use_container_width=True,type='primary'): st.session_state.selected_job=int(r['id']); st.rerun()
                else:
                    c.success(str(r.tecnico_esecutore or 'Eseguita')); c.caption(str(r.eseguita_il or ''))
                    if bool(r.semestrale): c.caption('☑ Semestrale')

st.caption('TESYX Manutenzioni v2 · Archivio separato · Dati condivisi online su Supabase.')
