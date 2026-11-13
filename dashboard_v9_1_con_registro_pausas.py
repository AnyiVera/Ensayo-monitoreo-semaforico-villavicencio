"""Dashboard V9 del prototipo IoT (SIMULACIÓN, sin sensores ni red física).
Ejecutar: python -m streamlit run dashboard_v9_iot.py
Conserva intactas las bases de V6 y V7; usa solo dashboard_v9_demo.db.
"""
from pathlib import Path
from io import BytesIO
from datetime import datetime, timezone
import json
import sqlite3
import time
import uuid
import shutil
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
DEMO = ROOT / 'dashboard_v9_demo.db'
NODES = {'SIM-01': 'La Rochela', 'SIM-02': 'Porfía'}
STEP = 5
SEND_EVERY = 30  # segundos: telemetría ordinaria, no muestreo
WARNING_AFTER = 5 * 60
SUSTAINED_AFTER = 60 * 60
BACKUP_EVERY = 24 * 3600
# 40 bytes x (24 h x 3600 / 30 s) = 115.200 bytes POR NODO
BUFFER_REFERENCE_BYTES = 115200

def now():
    return datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')

def parse(s):
    return datetime.fromisoformat(s)

def elapsed(s, end=None):
    if not s: return 0
    return max(0, ((parse(end) if end else datetime.now(timezone.utc).astimezone()) - parse(s)).total_seconds())

def clock(seconds):
    seconds = int(max(0, seconds))
    return f'{seconds//3600:02d}:{(seconds%3600)//60:02d}:{seconds%60:02d}'

def connect():
    con = sqlite3.connect(DEMO, timeout=10)
    con.row_factory = sqlite3.Row
    con.execute('PRAGMA busy_timeout=10000')
    return con

def init():
    with connect() as con:
        con.executescript('''
        CREATE TABLE IF NOT EXISTS config (node TEXT PRIMARY KEY, online INTEGER NOT NULL, disconnected_at TEXT, last_connected TEXT, last_sample TEXT, samples INTEGER NOT NULL DEFAULT 0);
        CREATE TABLE IF NOT EXISTS messages (id TEXT PRIMARY KEY, node TEXT NOT NULL, at TEXT NOT NULL, voltage REAL NOT NULL, state TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'Pendiente', sent_at TEXT);
        CREATE TABLE IF NOT EXISTS server (id TEXT PRIMARY KEY, node TEXT NOT NULL, received_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS transitions (id INTEGER PRIMARY KEY AUTOINCREMENT, node TEXT NOT NULL, event TEXT NOT NULL, at TEXT NOT NULL, duration_s REAL);
        CREATE TABLE IF NOT EXISTS power (node TEXT PRIMARY KEY, powered INTEGER NOT NULL DEFAULT 1, off_at TEXT);
        CREATE TABLE IF NOT EXISTS alerts (id TEXT PRIMARY KEY, node TEXT NOT NULL, kind TEXT NOT NULL, opened_at TEXT NOT NULL, confirmed_at TEXT, closed_at TEXT, duration_s REAL, status TEXT NOT NULL, details TEXT);
        CREATE TABLE IF NOT EXISTS detector (node TEXT PRIMARY KEY, previous_kind TEXT, streak INTEGER NOT NULL DEFAULT 0, active_id TEXT);
        CREATE TABLE IF NOT EXISTS gaps (id INTEGER PRIMARY KEY AUTOINCREMENT, node TEXT NOT NULL, started_at TEXT NOT NULL, ended_at TEXT NOT NULL, duration_s REAL NOT NULL, reason TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS runtime (key TEXT PRIMARY KEY, value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS monitor_pauses (id INTEGER PRIMARY KEY AUTOINCREMENT, last_seen TEXT NOT NULL, resumed_at TEXT NOT NULL, duration_s REAL NOT NULL, reason TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS reconnect_attempts (id INTEGER PRIMARY KEY AUTOINCREMENT, node TEXT NOT NULL, at TEXT NOT NULL, result TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS maintenance (id INTEGER PRIMARY KEY AUTOINCREMENT, alert_id TEXT NOT NULL, node TEXT NOT NULL, at TEXT NOT NULL, responsible TEXT NOT NULL, action TEXT NOT NULL, result TEXT NOT NULL, status TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS telemetry_extra (message_id TEXT PRIMARY KEY, current_a REAL, component_state TEXT, FOREIGN KEY(message_id) REFERENCES messages(id));
        
        ''')
        for node in NODES:
            con.execute('INSERT OR IGNORE INTO config(node,online,disconnected_at,last_connected,last_sample,samples) VALUES (?,1,NULL,?,NULL,0)', (node,now()))
            con.execute('INSERT OR IGNORE INTO power(node,powered,off_at) VALUES (?,1,NULL)',(node,))
            con.execute('INSERT OR IGNORE INTO detector(node,previous_kind,streak,active_id) VALUES (?,NULL,0,NULL)',(node,))

def open_alert(con,node,kind,detail='',confirmed=True):
    ident=uuid.uuid4().hex
    t=now()
    con.execute('INSERT INTO alerts(id,node,kind,opened_at,confirmed_at,status,details) VALUES (?,?,?,?,?,?,?)',
                (ident,node,kind,t,t if confirmed else None,'Activa',detail))
    return ident

def close_alert(con,node,kind):
    t=now()
    for r in con.execute("SELECT id,opened_at FROM alerts WHERE node=? AND kind=? AND status='Activa'",(node,kind)).fetchall():
        con.execute("UPDATE alerts SET closed_at=?,duration_s=?,status='Cerrada' WHERE id=?",(t,elapsed(r['opened_at'],t),r['id']))

def set_link(node, online):
    with connect() as con:
        row=con.execute('SELECT * FROM config WHERE node=?',(node,)).fetchone()
        if bool(row['online'])==online:return
        t=now()
        if online:
            con.execute('UPDATE config SET online=1,disconnected_at=NULL,last_connected=? WHERE node=?',(t,node))
            con.execute('INSERT INTO transitions(node,event,at,duration_s) VALUES (?,?,?,?)',(node,'Reconexión',t,elapsed(row['disconnected_at'],t)))
            for kind in ('Pérdida de comunicación','Advertencia local (5 min)','Falla de comunicación sostenida (1 h)'):
                close_alert(con,node,kind)
        else:
            con.execute('UPDATE config SET online=0,disconnected_at=? WHERE node=?',(t,node))
            con.execute('INSERT INTO transitions(node,event,at,duration_s) VALUES (?,?,?,NULL)',(node,'Desconexión',t))
            con.execute('DELETE FROM runtime WHERE key=?',('attempt_'+node,))
            open_alert(con,node,'Pérdida de comunicación','Enlace simulado desconectado')

def set_power(node,powered):
    with connect() as con:
        row=con.execute('SELECT * FROM power WHERE node=?',(node,)).fetchone()
        if bool(row['powered'])==powered:return
        t=now()
        if powered:
            con.execute('UPDATE power SET powered=1,off_at=NULL WHERE node=?',(node,))
            duration=elapsed(row['off_at'],t)
            con.execute('INSERT INTO gaps(node,started_at,ended_at,duration_s,reason) VALUES (?,?,?,?,?)',
                        (node,row['off_at'],t,duration,'Corte eléctrico simulado'))
            close_alert(con,node,'Sin energía / sin mediciones')
            # Reinicia confirmación: las lecturas antes del apagón no son consecutivas.
            con.execute('UPDATE detector SET previous_kind=NULL,streak=0,active_id=NULL WHERE node=?',(node,))
            # Si había alerta de voltaje abierta, se cierra al reanudar y se vuelve a evaluar.
            for kind in ('Sobretensión','Caída de tensión'):close_alert(con,node,kind)
        else:
            con.execute('UPDATE power SET powered=0,off_at=? WHERE node=?',(t,node))
            open_alert(con,node,'Sin energía / sin mediciones','Nodo apagado; no se generan lecturas')

def detect(con,node,kind):
    d=con.execute('SELECT * FROM detector WHERE node=?',(node,)).fetchone()
    previous=d['previous_kind']
    streak=d['streak']+1 if kind==previous else 1
    active=d['active_id']
    if kind!=previous and active:
        for old in ('Sobretensión','Caída de tensión'):
            close_alert(con,node,old)
        active=None
    if kind=='Normal':
        if active:
            for old in ('Sobretensión','Caída de tensión'):close_alert(con,node,old)
        active=None
    elif streak>=2 and not active:
        active=open_alert(con,node,kind,'Confirmada por dos lecturas consecutivas del mismo tipo')
    con.execute('UPDATE detector SET previous_kind=?,streak=?,active_id=? WHERE node=?',(kind,streak,active,node))

def voltage(node, n):
    phase = ((n-1)*STEP) % 180
    if node == 'SIM-01': return 140. if 35 <= phase < 65 else (95. if 115 <= phase < 140 else 120.)
    return 95. if 60 <= phase < 90 else (140. if 145 <= phase < 165 else 120.)

def sample():
    # La app solo adquiere datos cuando el proceso de Streamlit está activo.
    # Nunca inventa mediciones retroactivas después de una pausa.
    with connect() as con:
        for row in con.execute('SELECT * FROM config').fetchall():
            if not con.execute('SELECT powered FROM power WHERE node=?',(row['node'],)).fetchone()[0]:continue
            if row['last_sample'] and elapsed(row['last_sample']) < STEP: continue
            v = voltage(row['node'],row['samples']+1)
            state = 'Caída de tensión' if v < 110 else ('Sobretensión' if v > 130 else 'Normal')
            msg_id=uuid.uuid4().hex
            con.execute('INSERT INTO messages(id,node,at,voltage,state,status) VALUES (?,?,?,?,?,?)', (msg_id,row['node'],now(),v,state,'Pendiente'))
            # Variables complementarias simuladas; no son lecturas de sensores físicos.
            simulated_current = round(v / 240.0, 3)  # carga resistiva ficticia de 240 ohmios
            con.execute('INSERT INTO telemetry_extra(message_id,current_a,component_state) VALUES (?,?,?)',
                        (msg_id,simulated_current,'Operativo (simulado)' if v>=110 and v<=130 else 'Revisar (simulado)'))
            con.execute('UPDATE config SET samples=samples+1,last_sample=? WHERE node=?',(now(),row['node']))
            detect(con,row['node'],state)

def runtime_get(con, key):
    row = con.execute('SELECT value FROM runtime WHERE key=?', (key,)).fetchone()
    return row['value'] if row else None


def runtime_set(con, key, value):
    con.execute('INSERT OR REPLACE INTO runtime(key,value) VALUES (?,?)', (key,str(value)))


def monitor_heartbeat():
    """Detecta huecos entre actualizaciones; no infiere causa ni inventa muestras."""
    t = now()
    with connect() as con:
        previous = runtime_get(con, 'monitor_last_seen')
        if previous:
            gap = elapsed(previous, t)
            if gap >= 20:
                con.execute('INSERT INTO monitor_pauses(last_seen,resumed_at,duration_s,reason) VALUES (?,?,?,?)',
                            (previous,t,gap,'Sin actualización del panel; causa no determinada'))
        runtime_set(con, 'monitor_last_seen', t)


def escalation_and_retries():
    """Escalamiento y comprobación simulada del enlace; NO reconecta por sí solo."""
    with connect() as con:
        for row in con.execute('SELECT * FROM config WHERE online=0').fetchall():
            node = row['node']
            seconds = elapsed(row['disconnected_at'])
            for threshold, kind in ((WARNING_AFTER,'Advertencia local (5 min)'),
                                    (SUSTAINED_AFTER,'Falla de comunicación sostenida (1 h)')):
                if seconds >= threshold and not con.execute(
                    "SELECT 1 FROM alerts WHERE node=? AND kind=? AND status='Activa'",(node,kind)).fetchone():
                    open_alert(con,node,kind,f'Enlace simulado interrumpido durante al menos {threshold} segundos')
            key='attempt_'+node
            previous=runtime_get(con,key)
            if not previous or elapsed(previous)>=30:
                t=now()
                con.execute('INSERT INTO reconnect_attempts(node,at,result) VALUES (?,?,?)',
                            (node,t,'Sin enlace; comprobación simulada, sin recuperación automática'))
                runtime_set(con,key,t)


def send_pending():
    """Telemetría ordinaria cada 30 s; registros críticos confirmados con prioridad.

    El enlace es simulado y se restablece manualmente con el botón de reconexión.
    """
    with connect() as con:
        for row in con.execute('SELECT node FROM config WHERE online=1').fetchall():
            node=row['node']
            last=runtime_get(con,'sent_batch_'+node)
            batch_due=not last or elapsed(last)>=SEND_EVERY
            # Alertas de voltaje confirmadas, aún no transmitidas: prioridad inmediata.
            urgent = con.execute("""SELECT m.id FROM messages m
                WHERE m.node=? AND m.status='Pendiente' AND m.state!='Normal'
                AND EXISTS (SELECT 1 FROM alerts a WHERE a.node=m.node
                    AND a.kind=m.state AND a.confirmed_at IS NOT NULL
                    AND a.opened_at<=m.at AND (a.closed_at IS NULL OR a.closed_at>=m.at))
                ORDER BY m.at LIMIT 200""",(node,)).fetchall()
            ids=[r['id'] for r in urgent]
            if batch_due:
                others=con.execute("SELECT id FROM messages WHERE node=? AND status='Pendiente' ORDER BY at LIMIT 500",
                                   (node,)).fetchall()
                ids=list(dict.fromkeys(ids+[r['id'] for r in others]))
                runtime_set(con,'sent_batch_'+node,now())
            for ident in ids:
                msg=con.execute('SELECT id,node,at,voltage,state FROM messages WHERE id=?',(ident,)).fetchone()
                if not msg or msg['node'] not in NODES or not isinstance(msg['voltage'],(int,float)):
                    continue  # validación básica del formato y del nodo
                received=now()
                con.execute('INSERT OR IGNORE INTO server(id,node,received_at) VALUES (?,?,?)',(ident,node,received))
                con.execute("UPDATE messages SET status='Enviado',sent_at=? WHERE id=? AND status='Pendiente'",(received,ident))


def backup_if_due():
    """Copia SQLite consistente cada 24 h de actividad, sin tocar V6/V7."""
    with connect() as con:
        last=runtime_get(con,'last_backup')
        if last and elapsed(last)<BACKUP_EVERY:
            return None
    folder=ROOT/'respaldos_v91'
    folder.mkdir(exist_ok=True)
    target=folder/('dashboard_v9_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.db')
    source=connect()
    try:
        with sqlite3.connect(target) as dest:
            source.backup(dest)
    finally:
        source.close()
    with connect() as con:
        runtime_set(con,'last_backup',now())
    return target


def history_files():
    return sorted(list(ROOT.glob('ensayo_24h_*.db')) + list(ROOT.glob('ensayo_v7_*.db')), key=lambda p:p.name, reverse=True)

def history(db):
    uri = db.resolve().as_uri()+'?mode=ro'
    with sqlite3.connect(uri,uri=True,timeout=3) as con:
        names = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if 'mediciones' not in names: raise ValueError('No se encontró la tabla mediciones')
        count = con.execute('SELECT COUNT(*) FROM mediciones').fetchone()[0]
        per = pd.read_sql_query('SELECT nodo,COUNT(*) AS mediciones,MIN(fecha) AS primera,MAX(fecha) AS ultima FROM mediciones GROUP BY nodo',con)
        total_events = con.execute('SELECT COUNT(*) FROM eventos').fetchone()[0] if 'eventos' in names else 0
        total_incidents = con.execute('SELECT COUNT(*) FROM incidencias').fetchone()[0] if 'incidencias' in names else 0
        latest = pd.read_sql_query('SELECT fecha,nodo,voltaje,estado FROM mediciones ORDER BY fecha DESC LIMIT 150',con)
        interruptions = pd.read_sql_query("SELECT fecha,detalle FROM eventos WHERE tipo='Reanudación tras interrupción' ORDER BY fecha DESC",con) if 'eventos' in names and 'control' in names and con.execute("SELECT COUNT(*) FROM control WHERE clave='version' AND valor='7'").fetchone()[0] else pd.DataFrame()
        return count,per,total_events,total_incidents,latest,interruptions

def excel_datetime(value):
    """Convierte fechas ISO del registro en fechas Excel sin zona horaria."""
    if value is None or pd.isna(value) or value == '':
        return None
    try:
        return datetime.fromisoformat(str(value)).replace(tzinfo=None)
    except (ValueError, TypeError):
        return str(value)


def make_excel_report():
    """Genera un XLSX legible sin modificar los datos de SQLite."""
    with connect() as con:
        measurements = pd.read_sql_query(
            'SELECT m.id,m.node,m.at,m.voltage,e.current_a,e.component_state,m.state,m.status,m.sent_at FROM messages m LEFT JOIN telemetry_extra e ON e.message_id=m.id ORDER BY m.at', con)
        incidents = pd.read_sql_query(
            'SELECT id,node,kind,opened_at,confirmed_at,closed_at,duration_s,status,details '
            'FROM alerts ORDER BY opened_at', con)
        received = con.execute('SELECT COUNT(*) FROM server').fetchone()[0]
        rows = con.execute('SELECT node,online,samples FROM config ORDER BY node').fetchall()
        events = pd.read_sql_query('SELECT node,event,at,duration_s FROM transitions ORDER BY at',con)
        cuts = pd.read_sql_query('SELECT node,reason,started_at,ended_at,duration_s FROM gaps ORDER BY started_at',con)
        maintenance = pd.read_sql_query('SELECT node,at,responsible,action,result,status FROM maintenance ORDER BY at',con)

    wb = Workbook()
    ws_m = wb.active
    ws_m.title = 'Mediciones'
    ws_i = wb.create_sheet('Incidencias')
    ws_r = wb.create_sheet('Resumen general')
    ws_t = wb.create_sheet('Enlaces y cortes')
    ws_a = wb.create_sheet('Mantenimiento')

    def heading(ws, labels):
        ws.append(labels)
        ws.freeze_panes = 'A2'
        for cell in ws[1]:
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='17365D')
            cell.alignment = Alignment(vertical='center', wrap_text=True)
        ws.row_dimensions[1].height = 32

    def finish(ws, widths, date_columns=(), state_column=None):
        ws.auto_filter.ref = ws.dimensions
        for index, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(index)].width = width
        for row in ws.iter_rows(min_row=2):
            if row[0].row % 2 == 0:
                for cell in row:
                    cell.fill = PatternFill('solid', fgColor='F1F5F9')
            for idx in date_columns:
                cell = row[idx - 1]
                if isinstance(cell.value, datetime):
                    cell.number_format = 'dd/mm/yyyy hh:mm:ss'
            if state_column:
                cell = row[state_column - 1]
                val = str(cell.value or '').lower()
                if 'sobretensión' in val or 'caída' in val or 'sin energía' in val or 'pérdida' in val:
                    cell.font = Font(color='B42318', bold=True)
                elif val in ('normal', 'enviado', 'cerrada'):
                    cell.font = Font(color='16723A')
                elif val in ('pendiente', 'activa'):
                    cell.font = Font(color='A15C00', bold=True)

    heading(ws_m, ['ID de mensaje', 'Nodo', 'Ubicación', 'Fecha y hora',
                   'Voltaje (V)', 'Estado eléctrico', 'Estado de envío', 'Fecha de envío'])
    for rec in measurements.itertuples(index=False):
        ws_m.append([rec.id, rec.node, NODES.get(rec.node, ''), excel_datetime(rec.at),
                     rec.voltage, rec.current_a, rec.component_state, rec.state, rec.status, excel_datetime(rec.sent_at)])
    finish(ws_m, [36, 13, 20, 23, 15, 22, 24, 23, 19, 23], (4, 10), 8)
    for row in ws_m.iter_rows(min_row=2):
        row[4].number_format = '0.0'
        row[5].number_format = '0.000'

    heading(ws_i, ['ID de incidencia', 'Nodo', 'Ubicación', 'Tipo de alerta',
                   'Inicio', 'Confirmación', 'Cierre', 'Duración (s)', 'Estado', 'Detalle'])
    for rec in incidents.itertuples(index=False):
        ws_i.append([rec.id, rec.node, NODES.get(rec.node, ''), rec.kind,
                     excel_datetime(rec.opened_at), excel_datetime(rec.confirmed_at),
                     excel_datetime(rec.closed_at), rec.duration_s, rec.status, rec.details])
    finish(ws_i, [36, 13, 20, 30, 23, 23, 23, 18, 15, 65], (5, 6, 7), 4)
    for row in ws_i.iter_rows(min_row=2):
        row[7].number_format = '0.0'

    heading(ws_r, ['Indicador', 'SIM-01 · La Rochela', 'SIM-02 · Porfía', 'Total'])
    rowmap = {r['node']: r for r in rows}
    def node_count(node, condition=None):
        subset = measurements[measurements['node'] == node]
        return len(subset) if condition is None else int((subset['status'] == condition).sum())
    def alert_count(node, condition=None):
        subset = incidents[incidents['node'] == node]
        return len(subset) if condition is None else int((subset['status'] == condition).sum())
    for label, func in [
        ('Mediciones registradas', lambda n: node_count(n)),
        ('Mensajes enviados', lambda n: node_count(n, 'Enviado')),
        ('Mensajes pendientes', lambda n: node_count(n, 'Pendiente')),
        ('Incidencias registradas', lambda n: alert_count(n)),
        ('Incidencias activas', lambda n: alert_count(n, 'Activa')),
        ('Incidencias cerradas', lambda n: alert_count(n, 'Cerrada')),
    ]:
        x, y = func('SIM-01'), func('SIM-02')
        ws_r.append([label, x, y, x + y])
    ws_r.append(['Mensajes recibidos en servidor', None, None, received])
    ws_r.append(['Conexión simulada actual',
                 'Conectado' if rowmap['SIM-01']['online'] else 'Desconectado',
                 'Conectado' if rowmap['SIM-02']['online'] else 'Desconectado', None])
    ws_r.append(['Origen', 'Datos simulados de dashboard_v9_demo.db', None, None])
    ws_r.append(['Nota', 'El reporte es una fotografía de los datos al descargarlo.', None, None])
    finish(ws_r, [36, 42, 26, 20])

    heading(ws_t,['Nodo','Evento','Inicio','Fin','Duración (s)'])
    for rec in events.itertuples(index=False):
        ws_t.append([rec.node,rec.event,excel_datetime(rec.at),None,rec.duration_s])
    for rec in cuts.itertuples(index=False):
        ws_t.append([rec.node,rec.reason,excel_datetime(rec.started_at),excel_datetime(rec.ended_at),rec.duration_s])
    finish(ws_t,[15,38,24,24,20],(3,4))
    heading(ws_a,['Nodo','Fecha','Responsable','Acción','Resultado','Seguimiento'])
    for rec in maintenance.itertuples(index=False):
        ws_a.append([rec.node,excel_datetime(rec.at),rec.responsible,rec.action,rec.result,rec.status])
    finish(ws_a,[15,24,28,65,65,25],(2,))
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


st.set_page_config(page_title='Monitoreo semafórico IoT · V9',page_icon='🚦',layout='wide')
init()
st.title('🚦 Centro de monitoreo semafórico · V9.1')
st.caption('Villavicencio · SIM-01 La Rochela y SIM-02 Porfía · prototipo académico con voltajes y enlaces simulados')
st.info('El panel de demostración usa dashboard_v9_demo.db. Las bases de los ensayos V6/V7 se consultan en modo solo lectura y nunca se modifican.')

live, past = st.tabs(['📡 Monitoreo y enlace simulado','📚 Evidencia de ensayos V6/V7'])
with live:
    st.subheader('Control de comunicación y energía por nodo')
    configs = {}
    with connect() as con:
        for r in con.execute('SELECT * FROM config'): configs[r['node']] = dict(r)
    a,b = st.columns(2)
    for column,node in zip((a,b),NODES):
        with column:
            st.markdown(f'**{node} · {NODES[node]}**')
            st.write('🟢 Enlace conectado' if configs[node]['online'] else '🔴 Enlace desconectado')
            with connect() as con:
                power_on=bool(con.execute('SELECT powered FROM power WHERE node=?',(node,)).fetchone()[0])
            st.write('⚡ Energía disponible' if power_on else '🔴 Sin energía (simulado)')
            if st.button('Simular corte eléctrico' if power_on else 'Restablecer energía',key='power_'+node,use_container_width=True):
                set_power(node,not power_on)
                st.rerun()
            if st.button('Desconectar enlace' if configs[node]['online'] else 'Reconectar enlace',key='toggle_'+node,use_container_width=True):
                set_link(node,not bool(configs[node]['online']))
                st.rerun()
    st.caption('Sin enlace: se sigue midiendo y se acumulan pendientes. Sin energía: no se generan mediciones; los datos confirmados permanecen. Todo es simulado.')
    @st.fragment(run_every='2s')
    def live_panel():
        monitor_heartbeat()
        sample()
        escalation_and_retries()
        send_pending()
        backup_if_due()
        with connect() as con:
            rows = [dict(r) for r in con.execute('SELECT * FROM config ORDER BY node')]
            counts = {r['node']:(r['pending'],r['sent']) for r in con.execute("SELECT node,SUM(CASE WHEN status='Pendiente' THEN 1 ELSE 0 END) pending,SUM(CASE WHEN status='Enviado' THEN 1 ELSE 0 END) sent FROM messages GROUP BY node")}
            received = con.execute('SELECT COUNT(*) FROM server').fetchone()[0]
            transitions = pd.read_sql_query('SELECT node AS Nodo,event AS Evento,at AS Fecha,duration_s AS Duración_segundos FROM transitions ORDER BY id DESC LIMIT 30',con)
            df = pd.read_sql_query('SELECT at AS Fecha,node AS Nodo,voltage AS Voltaje,state AS Estado,status AS Envío FROM messages ORDER BY at DESC LIMIT 100',con)
            dup = con.execute('SELECT COUNT(*)-COUNT(DISTINCT id) FROM server').fetchone()[0]
            alerts=pd.read_sql_query('SELECT node AS Nodo,kind AS Tipo,opened_at AS Inicio,confirmed_at AS Confirmación,closed_at AS Fin,duration_s AS Duración_segundos,status AS Estado,details AS Detalle FROM alerts ORDER BY opened_at DESC LIMIT 100',con)
            gaps=pd.read_sql_query('SELECT node AS Nodo,started_at AS Inicio,ended_at AS Fin,duration_s AS Duración_segundos,reason AS Motivo FROM gaps ORDER BY id DESC LIMIT 50',con)
            power_states={r['node']:dict(r) for r in con.execute('SELECT * FROM power')}
            active_count=con.execute("SELECT COUNT(*) FROM alerts WHERE status='Activa'").fetchone()[0]
            pauses=pd.read_sql_query('SELECT last_seen AS Última_actividad,resumed_at AS Reanudación,duration_s AS Duración_segundos,reason AS Observación FROM monitor_pauses ORDER BY id DESC LIMIT 20',con)
        st.subheader('🖥️ Continuidad del sistema de monitoreo')
        if not pauses.empty:
            latest_pause=pauses.iloc[0]
            st.warning(f"Último intervalo sin actualización detectado: {clock(latest_pause['Duración_segundos'])}. Desde {latest_pause['Última_actividad']} hasta {latest_pause['Reanudación']}.")
            st.dataframe(pauses,hide_index=True,use_container_width=True)
        else:
            st.success('Sin interrupciones de actualización detectadas desde que se habilitó este registro.')
        st.caption('El intervalo se calcula a partir de la última actualización registrada. Puede deberse a apagado, suspensión, cierre del servidor o falta de sesiones activas. No equivale a tiempo sin energía de los nodos ni prueba que se monitoreó durante la pausa.')
        # Avisos visibles y actualizados en cada refresco del panel (aprox. 2 s).
        # La duración del enlace se obtiene de disconnected_at, no del número de ciclos.
        disconnected = [r for r in rows if not r['online']]
        unpowered = [r for r in rows if not power_states[r['node']]['powered']]
        if disconnected or unpowered:
            st.subheader('🚨 Avisos de disponibilidad en tiempo real')
            for r in disconnected:
                node = r['node']
                pending = counts.get(node, (0, 0))[0]
                seconds=elapsed(r['disconnected_at'])
                level=('FALLA SOSTENIDA · 1 H' if seconds>=SUSTAINED_AFTER
                       else 'ADVERTENCIA LOCAL · 5 MIN' if seconds>=WARNING_AFTER
                       else 'PÉRDIDA DE ENLACE')
                st.error(
                    f"📡 **{level} — SIN ENLACE — {node} · {NODES[node]}**  |  "
                    f"**Tiempo desconectado: {clock(elapsed(r['disconnected_at']))}**  |  "
                    f"**Mensajes pendientes: {pending}**  |  "
                    f"Desde: {r['disconnected_at']}"
                )
            for r in unpowered:
                node = r['node']
                off_at = power_states[node]['off_at']
                st.warning(
                    f"⚡ **SIN ENERGÍA SIMULADA — {node} · {NODES[node]}**  |  "
                    f"**Tiempo sin mediciones: {clock(elapsed(off_at))}**  |  "
                    f"Desde: {off_at}"
                )
            st.caption('Avisos de estados simulados. Si el computador se apaga, el panel no puede actualizarse hasta volver a iniciarse.')
        else:
            st.success('✅ Ambos nodos con enlace y energía simulados disponibles.')
        st.caption('Umbrales de enlace: aviso inmediato, advertencia local a los 5 min y falla sostenida a 1 h. Intentos de comprobación simulados cada 30 s; reconexión manual.')
        st.subheader('Estado actual')
        cols = st.columns(2)
        for col,row in zip(cols,rows):
            node = row['node']
            pending,sent = counts.get(node,(0,0))
            with col:
                st.markdown(f'**{node} · {NODES[node]}**')
                p=power_states[node]
                if not p['powered']:
                    st.error('Nodo sin energía (simulado)')
                    st.metric('Tiempo sin energía',clock(elapsed(p['off_at'])))
                else:
                    st.success('Energía disponible')
                if row['online']:
                    st.success('Comunicación disponible')
                    st.metric('Tiempo desconectado','00:00:00')
                else:
                    st.error('Sin comunicación con servidor (simulado)')
                    st.metric('Tiempo desconectado',clock(elapsed(row['disconnected_at'])))
                    st.caption('Inicio de desconexión: '+str(row['disconnected_at']))
                x,y = st.columns(2)
                x.metric('Pendientes',pending)
                y.metric('Enviados',sent)
                st.caption('Última muestra local: '+str(row['last_sample'] or 'Sin muestras'))
        c1,c2,c3 = st.columns(3)
        c1.metric('Mensajes recibidos en servidor',received)
        c2.metric('Mensajes pendientes',sum(v[0] for v in counts.values()))
        c3.metric('IDs duplicados en servidor',dup)
        st.metric('Alertas activas',active_count)
        st.subheader('🚨 Motor de alertas e incidencias')
        st.caption('Las alertas de voltaje requieren dos lecturas consecutivas anómalas del mismo tipo. Las de enlace y energía se activan al simular el cambio.')
        if not alerts.empty:st.dataframe(alerts,hide_index=True,use_container_width=True)
        else:st.info('Todavía no hay alertas.')
        st.subheader('⚡ Periodos sin mediciones por corte eléctrico')
        if not gaps.empty:st.dataframe(gaps,hide_index=True,use_container_width=True)
        else:st.caption('No hay interrupciones eléctricas finalizadas.')
        st.subheader('Mediciones simuladas recientes')
        if not df.empty:
            chart = df.sort_values('Fecha').pivot_table(index='Fecha',columns='Nodo',values='Voltaje',aggfunc='last')
            st.line_chart(chart)
            st.dataframe(df,hide_index=True,use_container_width=True)
        st.subheader('🔁 Comprobaciones simuladas de reconexión')
        with connect() as con:
            attempts=pd.read_sql_query('SELECT node AS Nodo,at AS Fecha,result AS Resultado FROM reconnect_attempts ORDER BY id DESC LIMIT 30',con)
        if not attempts.empty:st.dataframe(attempts,hide_index=True,use_container_width=True)
        st.subheader('Historial de desconexiones y reconexiones')
        if transitions.empty: st.caption('Todavía no hay cambios de enlace.')
        else: st.dataframe(transitions,hide_index=True,use_container_width=True)
        st.caption('Actualización automática aproximada cada 2 segundos. Muestreo nominal cada 5 segundos mientras el panel esté activo. No es un servicio independiente de adquisición 24/7.')
    live_panel()
    st.subheader('🛠️ Seguimiento de mantenimiento')
    with connect() as con:
        selectable=[dict(r) for r in con.execute('SELECT id,node,kind,status FROM alerts ORDER BY opened_at DESC LIMIT 300')]
        actions=pd.read_sql_query('SELECT node AS Nodo,at AS Fecha,responsible AS Responsable,action AS Acción,result AS Resultado,status AS Seguimiento FROM maintenance ORDER BY id DESC LIMIT 30',con)
    if selectable:
        with st.form('maintenance_form',clear_on_submit=True):
            selected=st.selectbox('Incidencia para seguimiento',selectable,format_func=lambda r:f"{r['node']} · {r['kind']} · {r['status']} · {r['id'][:8]}")
            responsible=st.text_input('Responsable')
            action=st.text_area('Acción de mantenimiento o verificación')
            result=st.text_area('Resultado observado')
            followup=st.selectbox('Estado del seguimiento',['Pendiente','En revisión','Intervención realizada'])
            if st.form_submit_button('Guardar seguimiento'):
                if responsible.strip() and action.strip() and result.strip():
                    with connect() as con:
                        con.execute('INSERT INTO maintenance(alert_id,node,at,responsible,action,result,status) VALUES (?,?,?,?,?,?,?)',
                                    (selected['id'],selected['node'],now(),responsible.strip(),action.strip(),result.strip(),followup))
                    st.success('Seguimiento registrado. La alerta técnica se cierra solo al normalizarse la condición.')
                else:st.warning('Completa responsable, acción y resultado.')
    if not actions.empty:st.dataframe(actions,hide_index=True,use_container_width=True)
    st.caption('El seguimiento humano no altera los sensores simulados ni cierra automáticamente alertas técnicas.')
    st.subheader('📊 Exportar reporte a Excel')
    st.caption('Incluye todas las mediciones, incidencias y un resumen por nodo. Las columnas se abren organizadas directamente en Excel.')
    st.info('Alcance: tensión, corriente y componentes son simulados; la corriente se estima con una carga ficticia de 240 Ω. El muestreo es de 5 s y el envío ordinario se agrupa cada 30 s. El presupuesto de 115.200 bytes/24 h corresponde a 40 bytes por registro cada 30 s POR NODO, no a una cuota física implementada en SQLite. Los respaldos se crean cada 24 h de ejecución del panel (sin servicio autónomo).')
    st.download_button(
        '📥 Descargar reporte completo (.xlsx)',
        data=make_excel_report(),
        file_name='reporte_monitoreo_v9.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )
with past:
    st.subheader('Consultar bases de datos históricas')
    files = history_files()
    if not files:
        st.warning('No hay bases ensayo_24h_*.db o ensayo_v7_*.db en esta carpeta. Copia el Dashboard junto a tus archivos de ensayo.')
    else:
        selection = st.selectbox('Selecciona un ensayo',files,format_func=lambda p:p.name)
        try:
            count,per,events,incidents,latest,interruptions = history(selection)
            x,y,z = st.columns(3)
            x.metric('Mediciones',count)
            y.metric('Eventos',events)
            z.metric('Incidencias',incidents)
            st.dataframe(per,hide_index=True,use_container_width=True)
            st.subheader('Últimas mediciones almacenadas')
            st.dataframe(latest,hide_index=True,use_container_width=True)
            if not interruptions.empty:
                st.subheader('Recuperaciones documentadas')
                st.dataframe(interruptions,hide_index=True,use_container_width=True)
            st.download_button('Descargar resumen por nodo',per.to_csv(index=False).encode('utf-8-sig'),f'{selection.stem}_nodos.csv','text/csv')
            st.caption('Los ensayos históricos no se convierten en datos en vivo; solo se consultan.')
        except (sqlite3.Error,ValueError,OSError) as exc:
            st.error(f'No se pudo consultar el ensayo: {exc}')
