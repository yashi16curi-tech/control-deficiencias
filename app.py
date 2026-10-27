import io
from html import escape
from PIL import Image, ImageOps, ImageDraw
import re
import zipfile
import unicodedata
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from docx import Document

st.set_page_config(page_title="Control de deficiencias", page_icon="◼", layout="wide", initial_sidebar_state="expanded")
ROOT = Path(__file__).resolve().parent
BASE = ROOT / "data" / "base_inicial.xlsm"
WORD = ROOT / "data" / "reporte_inicial.docx"
IMG_EXT = {'.jpg', '.jpeg', '.png', '.webp', '.bmp'}

st.markdown('''<style>
.stApp, [data-testid="stSidebar"] {background:#0c111b;color:#e8eef8}
.block-container{max-width:1320px;padding-top:1.4rem;padding-bottom:1rem}
h1,h2,h3{letter-spacing:-.035em} h1{font-size:2rem!important}
[data-testid="stMetric"]{background:#151d2b;border:1px solid #283447;border-radius:12px;padding:15px}
[data-testid="stMetricLabel"]{color:#9caec3} [data-testid="stMetricValue"]{color:#eff6ff}
[data-testid="stSidebar"]{border-right:1px solid #243044}
.stButton>button, [data-testid="stDownloadButton"] button{border-radius:10px;border:1px solid #273445;background:#151d2b;color:#e8eef8}
[data-testid="stDataFrame"]{border:1px solid #273445;border-radius:8px}
div[data-testid="stMetric"] > label {font-size:0.95rem}
.small-note{color:#94a3b8;font-size:.88rem}
.card-note{color:#93c5fd;font-size:.82rem;margin-top:-.3rem}

.ficha-tabla{border-collapse:collapse;width:100%;font-size:.91rem;table-layout:fixed;margin:0 0 1rem}
.ficha-tabla th,.ficha-tabla td{border:1px solid #526078;text-align:left;padding:10px 12px;vertical-align:top;word-break:break-word}
.ficha-tabla th{width:45%;font-weight:600;color:#bfdbfe;background:#182235}
.ficha-actividad{border:1px solid #526078;min-height:260px;display:flex;flex-direction:column;margin-bottom:1rem}
.ficha-titulo{font-size:.9rem;color:#bfdbfe;font-weight:600;padding:10px 12px 0}
.ficha-descripcion{padding:10px 12px 18px;white-space:pre-wrap;flex:1;font-size:.94rem;overflow-wrap:anywhere}
.ficha-metadata{display:flex;flex-wrap:wrap;border-top:1px solid #526078;font-size:.85rem}
.ficha-metadata span{padding:9px 12px;border-right:1px solid #526078;flex:1;min-width:140px}
.ficha-separacion{height:15px}
.ficha-foto-titulo{border:1px solid #34a853;border-radius:2px;text-align:center;padding:8px 4px;margin-bottom:12px;font-weight:600;color:#dcfce7;font-size:.92rem}
</style>''', unsafe_allow_html=True)


def norm(v):
    s = unicodedata.normalize('NFKD', str(v).strip())
    return ''.join(c for c in s if not unicodedata.combining(c)).upper()


def ident(v):
    if pd.isna(v):
        return ''
    return re.sub(r'\.0$', '', str(v).strip())


@st.cache_data(show_spinner=False)
def excel_df(raw):
    sheets = pd.read_excel(io.BytesIO(raw), sheet_name=None, engine='openpyxl', dtype=object)
    sheets = {norm(k): v for k, v in sheets.items()}
    if 'CUADRO 1' not in sheets or 'CUADRO 2' not in sheets:
        raise ValueError('Se necesitan las hojas CUADRO 1 y CUADRO 2.')
    a = sheets['CUADRO 1'].copy()
    b = sheets['CUADRO 2'].copy()
    a.columns = [norm(k) for k in a.columns]
    b.columns = [norm(k) for k in b.columns]
    for field in ['N°', 'CODIGO DEF. MAXIMUS', 'ALIMENTADOR', 'TIPO DE DEFICIENCIA']:
        if field not in a:
            raise ValueError(f'Falta la columna {field} de CUADRO 1.')
    a = a[a['CODIGO DEF. MAXIMUS'].notna()].copy()
    b = b[b['N°'].notna()].copy()
    a['_id'] = a['N°'].map(ident)
    b['_id'] = b['N°'].map(ident)
    if a['_id'].duplicated().any() or b['_id'].duplicated().any():
        raise ValueError('La columna N° tiene duplicados; no se puede unir de manera segura.')
    d = a.merge(b.drop(columns=['N°']), on='_id', how='left', validate='one_to_one')
    mapping = {
        'N°': 'N°', 'CODIGO DEF. MAXIMUS': 'Maximus', 'UUNN': 'Unidad', 'CODIGO POSTE / SED': 'Poste / SED',
        'ALIMENTADOR': 'Alimentador', 'TIPO DE DEFICIENCIA': 'Tipificación',
        'DESCRIPCION DE TIPIFICACION': 'Descripción', 'FECHA': 'Fecha',
        'ESTADO DE TIPIFICACION': 'Estado Excel', 'DESCRIPCION DE LAS ACTIVIDADES REALIZADAS': 'Actividad',
        'N°OM': 'OM', 'ING. QUE REPORTA': 'Responsable', 'REF': 'Referencia'
    }
    d = d.rename(columns=mapping)
    for name in mapping.values():
        if name not in d:
            d[name] = ''
    for k in ['Maximus', 'Unidad', 'Poste / SED', 'Alimentador', 'Tipificación', 'Estado Excel', 'OM', 'Responsable']:
        d[k] = d[k].map(ident)
    for k in ['Descripción', 'Actividad', 'Referencia']:
        d[k] = d[k].fillna('').astype(str).str.strip()
    d['Fecha'] = pd.to_datetime(d['Fecha'], errors='coerce')
    return d


@st.cache_data(show_spinner=False)
def word_counts(raw):
    doc = Document(io.BytesIO(raw))
    found = {}
    active = None
    for el in doc.element.body.iterchildren():
        if el.tag.endswith('}tbl'):
            m = re.search(r'\b77\d{10}\b', ' '.join(el.itertext()))
            if m:
                active = m.group()
                found[active] = 0
        if active:
            found[active] += len(el.xpath('.//*[local-name()="blip"]'))
    return found


def local_photos(base):
    result = {}
    if not base.is_dir():
        return result
    for folder in base.iterdir():
        if not folder.is_dir() or not re.fullmatch(r'\d{12}', folder.name.strip()):
            continue
        result[folder.name.strip()] = sorted(
            [p for p in folder.rglob('*') if p.is_file() and p.suffix.lower() in IMG_EXT],
            key=lambda x: x.name.casefold()
        )
    return result


@st.cache_data(show_spinner=False)
def zip_photos(raw):
    result = {}
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        for item in z.infolist():
            if item.is_dir() or item.file_size > 20_000_000:
                continue
            parts = Path(item.filename.replace('\\', '/')).parts
            if '..' in parts or not parts:
                continue
            idx = next((i for i, p in enumerate(parts[:-1]) if re.fullmatch(r'\d{12}', p)), None)
            if idx is None or Path(parts[-1]).suffix.lower() not in IMG_EXT:
                continue
            code = parts[idx]
            result.setdefault(code, []).append((parts[-1], z.read(item)))
    for code in result:
        result[code].sort(key=lambda pair: pair[0].lower())
    return result


def image_order(items):
    def key(item):
        name = norm(item[0])
        if any(w in name for w in ('ANTES', 'ANTES DE', 'BEFORE', 'INICIAL')):
            return (0, name)
        if any(w in name for w in ('DESPUES', 'DESPUÉS', 'AFTER', 'FINAL')):
            return (2, name)
        return (1, name)
    return sorted(items, key=key)


def photo_slots(items):
    """Separate original, detail and final images by names, never by file count."""
    ordered = image_order(items)
    before = next((it for it in ordered if any(t in norm(it[0]) for t in ('FOTO_ANTES', 'FOTO ANTES', 'ANTES', 'INICIAL')) and not any(t in norm(it[0]) for t in ('DETALLE', 'ACERCAMIENTO'))), None)
    detail = next((it for it in ordered if any(t in norm(it[0]) for t in ('DETALLE_ANTES', 'DETALLE ANTES', 'ACERCAMIENTO_ANTES'))), None)
    after = next((it for it in ordered if any(t in norm(it[0]) for t in ('DESPUES', 'AFTER', 'POSTERIOR', 'FINAL')) and 'DETALLE' not in norm(it[0])), None)
    if before is None:
        # Only choose unlabelled files as main 'before' when no explicit label is present.
        before = next((it for it in ordered if it not in (detail, after) and not any(t in norm(it[0]) for t in ('DETALLE', 'DESPUES', 'POSTERIOR', 'FINAL'))), None)
    return before, detail, after


def open_photo(item):
    if item is None:
        return None
    _, source = item
    try:
        if isinstance(source, Path):
            with Image.open(source) as im:
                return ImageOps.exif_transpose(im).convert('RGB')
        with Image.open(io.BytesIO(source)) as im:
            return ImageOps.exif_transpose(im).convert('RGB')
    except Exception:
        return None


def photo_frame(item, detail=None):
    """Portrait before/after boxes following the Word layout; circle detail over before."""
    width, height = 720, 870
    frame = Image.new('RGB', (width, height), '#121826')
    original = open_photo(item)
    if original is not None:
        original.thumbnail((width - 20, height - 20), Image.Resampling.LANCZOS)
        frame.paste(original, ((width - original.width) // 2, (height - original.height) // 2))
    detail_image = open_photo(detail)
    if original is not None and detail_image is not None:
        diameter = 280
        detail_image = ImageOps.fit(detail_image, (diameter, diameter), method=Image.Resampling.LANCZOS)
        mask = Image.new('L', (diameter, diameter), 0)
        drawmask = ImageDraw.Draw(mask)
        drawmask.ellipse((2, 2, diameter-3, diameter-3), fill=255)
        # The detail is an inset inside the BEFORE photo, as in the Word report.
        x, y = 100, 385
        frame.paste(detail_image, (x, y), mask)
        ImageDraw.Draw(frame).ellipse((x, y, x+diameter-1, y+diameter-1), outline='#f8fafc', width=4)
    ImageDraw.Draw(frame).rectangle((1, 1, width-2, height-2), outline='#738095', width=2)
    return frame


def clean(value):
    return escape(str(value if pd.notna(value) and str(value).strip() else '—'))


st.title('Seguimiento de deficiencias')
st.caption('Control documental · Huancayo')

with st.sidebar:
    st.subheader('Archivos')
    exc = st.file_uploader('Excel actualizado', type=['xlsx', 'xlsm'])
    doc = st.file_uploader('Reporte Word actualizado', type=['docx'])
    st.divider()
    st.subheader('Fotografías')
    folder = st.text_input(
        'Ruta de INSPECCIONES',
        value=str(ROOT / 'INSPECCIONES'),
        help='Usa una ruta absoluta o relativa. Si la carpeta está fuera del proyecto, por ejemplo ..\\INSPECCIONES o C:\\Users\\...\\INSPECCIONES'
    )
    photos_zip = st.file_uploader('O cargar INSPECCIONES.zip', type=['zip'])
    st.caption('La carpeta o el ZIP se usan solo para visualizar fotos. El estado se calcula con el Word.')

try:
    df = excel_df(exc.getvalue() if exc else BASE.read_bytes())
    evidence = word_counts(doc.getvalue() if doc else WORD.read_bytes())
    if photos_zip:
        photo_map = zip_photos(photos_zip.getvalue())
    else:
        photo_map = {k: [(p.name, p) for p in v] for k, v in local_photos(Path(folder).expanduser()).items()}
except Exception as e:
    st.error(f'No se pudo abrir la información: {e}')
    st.stop()


def classification(code):
    n = evidence.get(code)
    if n is None:
        return 'Sin evidencia'
    return 'Cerrada' if n >= 4 else 'Pendiente'


df['Seguimiento'] = df['Maximus'].map(classification)
df['Fotos carpeta'] = df['Maximus'].map(lambda c: len(photo_map.get(c, [])))
df['Fotos Word'] = df['Maximus'].map(lambda c: max(0, evidence.get(c, 0) // 2))

# filtros dependientes
with st.sidebar:
    st.divider()
    st.subheader('Filtros')
    filtered = df.copy()
    for column in ['Unidad', 'Alimentador', 'Tipificación']:
        opts = sorted(v for v in filtered[column].unique() if v)
        selected = st.selectbox(column, ['Todos'] + opts, key='filter_' + column)
        if selected != 'Todos':
            filtered = filtered[filtered[column] == selected]

    # el filtro de seguimiento no se aplica de inmediato si el usuario usa los botones KPI
    seg_options = ['Todos'] + [v for v in ['Cerrada', 'Pendiente', 'Sin evidencia'] if v in set(filtered['Seguimiento'])]
    seg_selected = st.selectbox('Seguimiento', seg_options, key='filter_Seguimiento')
    if seg_selected != 'Todos':
        filtered = filtered[filtered['Seguimiento'] == seg_selected]

    q = st.text_input('Buscar Maximus / poste / OM', placeholder='Código o texto')
    if q.strip():
        terms = ['Maximus', 'Poste / SED', 'OM', 'Descripción']
        mask = pd.Series(False, index=filtered.index)
        for c in terms:
            mask |= filtered[c].astype(str).str.contains(q.strip(), case=False, regex=False, na=False)
        filtered = filtered[mask]

# vista rápida desde KPI
if 'quick_state' not in st.session_state:
    st.session_state['quick_state'] = 'Todos'
if 'selected_code' not in st.session_state:
    st.session_state['selected_code'] = None

# si el sidebar ya filtró por seguimiento, ese tiene prioridad visible
active_state = seg_selected if seg_selected != 'Todos' else st.session_state['quick_state']
view = filtered.copy()
if active_state != 'Todos':
    view = view[view['Seguimiento'] == active_state]

n_total = len(filtered)
closed_total = int(filtered['Seguimiento'].eq('Cerrada').sum())
pending_total = int(filtered['Seguimiento'].eq('Pendiente').sum())
unknown_total = int(filtered['Seguimiento'].eq('Sin evidencia').sum())

k1, k2, k3, k4 = st.columns(4)
with k1:
    st.metric('Total', n_total)
    if st.button('Ver todo', key='btn_total', use_container_width=True):
        st.session_state['quick_state'] = 'Todos'
        st.rerun()
with k2:
    st.metric('Cerradas', closed_total)
    if st.button('Ver cerradas', key='btn_cerradas', use_container_width=True):
        st.session_state['quick_state'] = 'Cerrada'
        st.rerun()
with k3:
    st.metric('Pendientes', pending_total)
    if st.button('Ver pendientes', key='btn_pendientes', use_container_width=True):
        st.session_state['quick_state'] = 'Pendiente'
        st.rerun()
with k4:
    cumplimiento = f'{100*closed_total/n_total:.1f}%' if n_total else '—'
    st.metric('Cumplimiento', cumplimiento)
    if st.button('Sin evidencia', key='btn_sin_evidencia', use_container_width=True):
        st.session_state['quick_state'] = 'Sin evidencia'
        st.rerun()

st.progress(closed_total / n_total if n_total else 0)
state_label = 'Todos los registros' if active_state == 'Todos' else f'Vista actual: {active_state}'
st.markdown(f"<div class='small-note'>{state_label}</div>", unsafe_allow_html=True)

left, right = st.columns([1.15, 1], gap='large')
with left:
    st.subheader('Estado de atención')
    count = filtered['Seguimiento'].value_counts().reindex(['Cerrada', 'Pendiente', 'Sin evidencia'], fill_value=0).reset_index()
    count.columns = ['Estado', 'Cantidad']
    if n_total:
        fig = px.bar(
            count, x='Estado', y='Cantidad', color='Estado', text='Cantidad',
            color_discrete_map={'Cerrada': '#38bdf8', 'Pendiente': '#f59e0b', 'Sin evidencia': '#64748b'}
        )
        fig.update_layout(
            template='plotly_dark', paper_bgcolor='#0c111b', plot_bgcolor='#0c111b',
            height=300, margin=dict(l=0, r=0, t=10, b=0), showlegend=False,
            xaxis_title='', yaxis_title=''
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info('No hay registros con estos filtros.')

with right:
    st.subheader('Por alimentador')
    if n_total:
        grp = filtered.groupby('Alimentador')['Seguimiento'].agg(
            Total='size',
            Cerradas=lambda x: (x == 'Cerrada').sum()
        ).reset_index()
        grp['Pendientes'] = grp['Total'] - grp['Cerradas']
        grp['Cumplimiento %'] = (100 * grp['Cerradas'] / grp['Total']).round(1)
        grp = grp.sort_values(['Cumplimiento %', 'Total'], ascending=[False, False])
        grp['Etiqueta'] = grp['Cerradas'].astype(str) + '/' + grp['Total'].astype(str)
        fig2 = px.bar(
            grp, y='Alimentador', x='Cumplimiento %', orientation='h', text='Etiqueta',
            color='Cumplimiento %', color_continuous_scale=['#1d4ed8', '#38bdf8'], range_x=[0, 100]
        )
        fig2.update_layout(
            template='plotly_dark', paper_bgcolor='#0c111b', plot_bgcolor='#0c111b',
            height=300, margin=dict(l=0, r=0, t=10, b=0), coloraxis_showscale=False,
            xaxis_title='Cumplimiento %', yaxis_title=''
        )
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info('Sin datos.')

list_cols = ['Maximus', 'Poste / SED', 'Alimentador', 'Tipificación', 'OM', 'Estado Excel', 'Seguimiento', 'Fotos Word', 'Fotos carpeta']
if active_state != 'Todos':
    st.subheader(f'Registros · {active_state}')
    st.dataframe(view[list_cols], hide_index=True, use_container_width=True, height=240)

selected_candidates = view['Maximus'].dropna().drop_duplicates().tolist()
if selected_candidates:
    default_index = 0
    if st.session_state['selected_code'] in selected_candidates:
        default_index = selected_candidates.index(st.session_state['selected_code'])
    selection = st.selectbox(
        'Selecciona un registro',
        selected_candidates,
        index=default_index,
        format_func=lambda c: f'{c}  ·  {view.loc[view.Maximus == c, "Poste / SED"].iloc[0]}'
    )
    st.session_state['selected_code'] = selection
else:
    selection = None
    st.info('No hay registros disponibles en la vista actual.')

if selection:
    st.subheader('Ficha de intervención')
    rec = df.loc[df['Maximus'] == selection].iloc[0]

    # The two compact tables mirror the Word's identification and activity sections.
    details, activity = st.columns([0.92, 1.35], gap='medium')
    with details:
        rows = [
            ('UUNN', rec['Unidad']),
            ('Código de poste', rec['Poste / SED']),
            ('Código def. Maximus', rec['Maximus']),
            ('Tipificación', rec['Tipificación']),
            ('Fecha', rec['Fecha'].strftime('%d/%m/%Y') if pd.notna(rec['Fecha']) else '—'),
            ('Estado de tipificación', rec['Estado Excel']),
        ]
        html_rows = ''.join(
            f'<tr><th>{clean(title)}</th><td>{clean(value)}</td></tr>' for title, value in rows
        )
        st.markdown(f'<table class="ficha-tabla">{html_rows}</table>', unsafe_allow_html=True)
    with activity:
        st.markdown(f"""<div class='ficha-actividad'>
            <div class='ficha-titulo'>Descripción de las actividades realizadas</div>
            <div class='ficha-descripcion'>{clean(rec['Actividad'])}</div>
            <div class='ficha-metadata'>
                <span><b>N° OM:</b> {clean(rec['OM'])}</span>
                <span><b>Ing. que reporta:</b> {clean(rec['Responsable'])}</span>
                <span><b>Seguimiento:</b> {clean(rec['Seguimiento'])}</span>
            </div>
        </div>""", unsafe_allow_html=True)

    pictures = image_order(photo_map.get(selection, []))
    before, detail_before, after = photo_slots(pictures)
    st.markdown('<div class="ficha-separacion"></div>', unsafe_allow_html=True)
    first, second = st.columns(2, gap='medium')
    with first:
        st.markdown('<div class="ficha-foto-titulo">Foto antes de intervenir la deficiencia</div>', unsafe_allow_html=True)
        st.image(photo_frame(before, detail_before), use_container_width=True)
    with second:
        st.markdown('<div class="ficha-foto-titulo">Foto después de intervenir la deficiencia</div>', unsafe_allow_html=True)
        st.image(photo_frame(after), use_container_width=True)
    if not pictures:
        st.caption('No se encontraron imágenes en INSPECCIONES para este código. Verifica la ruta o carga el ZIP.')
    elif before is None or after is None:
        st.caption('Las fotos de detalle pertenecen al ANTES y no se utilizan como fotografía posterior. Los recuadros sin foto se mantienen vacíos.')

st.download_button(
    'Exportar CSV',
    view[list_cols + ['Descripción', 'Actividad', 'Responsable']].to_csv(index=False).encode('utf-8-sig'),
    file_name='seguimiento_filtrado.csv',
    mime='text/csv'
)

st.caption('Cierre documental según el Word. Foto_antes y Detalle_antes pertenecen a la misma evidencia; no equivalen a una foto después. INSPECCIONES sirve para visualizar imágenes.')
