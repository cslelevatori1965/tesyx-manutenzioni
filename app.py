import streamlit as st
import base64
import json
from supabase import create_client

st.set_page_config(page_title='Collaudo Supabase')
st.title('Collaudo Supabase')
password = st.secrets.get('ADMIN_PASSWORD')
if not password or st.text_input('Password amministratore', type='password') != password:
    st.stop()
url = str(st.secrets.get('SUPABASE_URL', '')).rstrip('/')
key = st.secrets.get('SUPABASE_KEY') or st.secrets.get('SUPABASE_SERVICE_KEY')
if not key or url != 'https://bccexizxccaymwpgzsru.supabase.co':
    st.error('Configurazione mancante o progetto diverso da quello previsto.')
    st.stop()
role = 'non determinabile localmente'
try:
    part = key.split('.')[1]
    role = json.loads(base64.urlsafe_b64decode(part + '=' * (-len(part) % 4))).get('role', role)
except Exception:
    if key.startswith('sb_secret_'): role = 'chiave secret moderna'
    elif key.startswith('sb_publishable_'): role = 'chiave publishable moderna'
st.write('Ruolo dichiarato dalla chiave:', role)
if role == 'anon': st.warning('La chiave effettivamente letta dal programma è anon.')
if st.button('Esegui collaudo'):
    try:
        client = create_client(url, key)
    except Exception as exc:
        st.error('Inizializzazione fallita: ' + type(exc).__name__)
        st.stop()
    for table in ['impianti_v4','tecnici_v4','programmazione_v4','manutenzioni_v4','stati_impianti_v4','storico_annullamenti_v4','storico_programmazione_v4']:
        try:
            result = client.table(table).select('*', count='exact').limit(1).execute()
            st.success(f'{table}: lettura OK — {result.count} righe')
        except Exception as exc:
            code = str(getattr(exc, 'code', '') or '')
            descriptions = {'42501':'permessi insufficienti','PGRST205':'tabella non trovata nella cache API','42P01':'tabella inesistente','PGRST301':'token non valido'}
            st.error(f'{table}: {type(exc).__name__} {code} — {descriptions.get(code, "accesso non riuscito")}')
    st.info('Test in sola lettura. Nessun dato modificato. I permessi di scrittura non sono ancora verificati.')
