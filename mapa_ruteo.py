"""Mapa y ruteo opcionales. No inventa ubicaciones de postes."""
import io
import html
import math
import unicodedata
from pathlib import Path

import folium
import pandas as pd
import requests
import streamlit as st
from streamlit_folium import st_folium
from pyproj import Transformer


def _norm(text):
    raw = unicodedata.normalize('NFKD', str(text).strip())
    return ''.join(c for c in raw if not unicodedata.combining(c)).lower().replace(' ', '').replace('_', '')


def _code(v):
    if pd.isna(v):
        return ''
    return str(v).strip().removesuffix('.0')


def parse_coordinates(raw, suffix):
    if suffix.lower() == '.csv':
        try:
            d = pd.read_csv(io.BytesIO(raw), sep=None, engine='python', dtype=str, encoding='utf-8-sig')
        except UnicodeDecodeError:
            d = pd.read_csv(io.BytesIO(raw), sep=None, engine='python', dtype=str, encoding='latin1')
    else:
        d = pd.read_excel(io.BytesIO(raw), dtype=str, engine='openpyxl')
    names = {_norm(c): c for c in d.columns}
    def pick(*options):
        return next((names[v] for v in options if v in names), None)
    code_col = pick('maximus', 'codigomaximus', 'codigodefmaximus', 'codigodef.maximus', 'solicitud')
    lat_col = pick('latitud', 'latitude', 'lat')
    lon_col = pick('longitud', 'longitude', 'lon', 'lng')
    # El Excel del usuario utiliza UTM WGS84 zona 18S, EPSG:32718.
    east_col = pick('xeste', 'este', 'easting')
    north_col = pick('ynorte', 'norte', 'northing')
    epsg_col = pick('epsg', 'sistemadereferencia', 'crs')
    if code_col and east_col and north_col and not (lat_col and lon_col):
        res = d[[code_col, east_col, north_col]].copy()
        res.columns = ['Maximus', 'X_ESTE', 'Y_NORTE']
        res['Maximus'] = res['Maximus'].map(_code)
        for k in ('X_ESTE', 'Y_NORTE'):
            res[k] = pd.to_numeric(res[k].astype(str).str.strip().str.replace(',', '.', regex=False), errors='coerce')
        if epsg_col:
            epsgs = set(d.loc[res['X_ESTE'].notna() & res['Y_NORTE'].notna(), epsg_col].dropna().astype(str).str.strip().str.upper())
            if epsgs != {'EPSG:32718'}:
                raise ValueError(f'El archivo tiene sistemas EPSG no soportados: {sorted(epsgs)}. Se requiere EPSG:32718.')
        res = res.dropna(subset=['X_ESTE', 'Y_NORTE'])
        transformer = Transformer.from_crs('EPSG:32718', 'EPSG:4326', always_xy=True)
        lon, lat = transformer.transform(res['X_ESTE'].to_numpy(), res['Y_NORTE'].to_numpy())
        res['Latitud'], res['Longitud'] = lat, lon
        res = res[['Maximus', 'Latitud', 'Longitud']]
    else:
        if not all((code_col, lat_col, lon_col)):
            raise ValueError('Se requiere SOLICITUD, X_ESTE, Y_NORTE, EPSG:32718; o Maximus, Latitud, Longitud.')
        res = d[[code_col, lat_col, lon_col]].copy()
        res.columns = ['Maximus', 'Latitud', 'Longitud']
        res['Maximus'] = res['Maximus'].map(_code)
    # Validaciones comunes para ambos formatos.
    for c in ['Latitud', 'Longitud']:
        res[c] = pd.to_numeric(res[c].astype(str).str.strip().str.replace(',', '.', regex=False), errors='coerce')
    res = res.dropna(subset=['Latitud', 'Longitud'])
    res = res[res['Latitud'].between(-90, 90) & res['Longitud'].between(-180, 180)]
    if res['Maximus'].duplicated().any():
        raise ValueError('El archivo contiene códigos Maximus duplicados con coordenadas. Cada código debe tener una ubicación.')
    return res


def dist_km(a, b):
    p1, p2 = math.radians(a['Latitud']), math.radians(b['Latitud'])
    dp = p2 - p1
    dl = math.radians(b['Longitud'] - a['Longitud'])
    h = math.sin(dp/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 6371 * 2 * math.asin(min(1, math.sqrt(h)))


def nearby_order(d, start_code):
    remaining = d.to_dict('records')
    current = next((r for r in remaining if r['Maximus'] == start_code), remaining[0])
    ordered = [current]
    remaining.remove(current)
    while remaining:
        nxt = min(remaining, key=lambda r: dist_km(current, r))
        ordered.append(nxt)
        remaining.remove(nxt)
        current = nxt
    return ordered


@st.cache_data(show_spinner=False, ttl=600)
def road_route(points):
    coords = ';'.join(f'{lon:.7f},{lat:.7f}' for lat,lon in points)
    url = f'https://router.project-osrm.org/route/v1/driving/{coords}'
    try:
        r = requests.get(url, params={'overview': 'full', 'geometries': 'geojson', 'steps': 'false'}, timeout=13)
        r.raise_for_status()
        body = r.json()
        if body.get('code') != 'Ok' or not body.get('routes'):
            return None
        route = body['routes'][0]
        return {'coordinates': route['geometry']['coordinates'], 'distance_km': route['distance']/1000}
    except (requests.RequestException, ValueError, KeyError):
        return None


def show_map(df, project_dir):
    st.header('Mapa y ruteo')
    st.caption('Haz clic en una burbuja para ver el código de poste, Maximus y su estado.')
    default = Path(project_dir) / 'data' / 'coordenadas_puntos.xlsx'
    uploaded = st.file_uploader('Coordenadas de postes (CSV o Excel)', type=['csv','xlsx'], key='map_coords')
    if uploaded:
        raw, suffix = uploaded.getvalue(), Path(uploaded.name).suffix
    elif default.is_file():
        raw, suffix = default.read_bytes(), '.xlsx'
    else:
        raw, suffix = None, None
    if raw is None:
        st.info('Carga el Excel de coordenadas con SOLICITUD, X_ESTE, Y_NORTE, EPSG:32718, o un archivo con Maximus, Latitud y Longitud.')
        sample = df[['Maximus','Poste / SED','Alimentador']].drop_duplicates('Maximus').copy()
        sample['Latitud'] = ''
        sample['Longitud'] = ''
        st.download_button('Descargar plantilla de coordenadas', sample.to_csv(index=False).encode('utf-8-sig'), 'plantilla_ubicaciones.csv', 'text/csv')
        return
    try:
        pos = parse_coordinates(raw, suffix)
    except Exception as exc:
        st.error(f'No se pudieron leer las coordenadas: {exc}')
        return
    points = df.merge(pos, on='Maximus', how='inner')
    if points.empty:
        st.warning('No hay códigos Maximus coincidentes entre los filtros actuales y las coordenadas cargadas.')
        return
    st.caption(f'{len(points)} postes con ubicación · {len(df)-len(points)} sin ubicación en la selección actual')
    if len(points)>250:
        st.warning('Mostrando los primeros 250 postes para mantener el mapa fluido.')
        points=points.head(250)
    col1, col2 = st.columns([1.3,1])
    with col1:
        status = st.multiselect('Mostrar estados', sorted(points['Seguimiento'].dropna().unique()), default=sorted(points['Seguimiento'].dropna().unique()))
    points = points[points['Seguimiento'].isin(status)].copy()
    if points.empty:
        st.info('No hay postes para estos estados.')
        return
    with col2:
        do_route = st.toggle('Mostrar ruta por carretera', value=False)
    ordered = None
    route = None
    if do_route:
        if len(points)>25:
            st.warning('El ruteo usa hasta 25 puntos. Filtra primero por alimentador o estado; los marcadores siguen disponibles.')
        elif len(points)<2:
            st.info('Se necesitan al menos 2 postes geolocalizados para trazar una ruta.')
        else:
            start = st.selectbox('Poste inicial', points['Maximus'].tolist(), format_func=lambda c: f'{points.loc[points.Maximus==c,"Poste / SED"].iloc[0]} · {c}')
            ordered = nearby_order(points, start)
            route = road_route(tuple((r['Latitud'],r['Longitud']) for r in ordered))
            if route is None:
                st.warning('No se pudo obtener la ruta de carretera. Se muestran solo los marcadores; no se dibuja una ruta recta que pueda confundirse con una vía real.')
            else:
                st.caption(f"Ruta por carretera: {route['distance_km']:.1f} km (orden aproximado por cercanía; no garantiza el óptimo).")
    center = [points['Latitud'].mean(), points['Longitud'].mean()]
    m = folium.Map(location=center, zoom_start=12, tiles='OpenStreetMap', control_scale=True, prefer_canvas=True)
    if route:
        line = [(lat, lon) for lon,lat in route['coordinates']]
        folium.PolyLine(line, color='#2563eb', weight=5, opacity=.9).add_to(m)
    group = folium.FeatureGroup(name='Postes')
    for _,row in points.iterrows():
        poste = html.escape(str(row['Poste / SED']))
        code = html.escape(str(row['Maximus']))
        status_label = html.escape(str(row['Seguimiento']))
        color = '#16a34a' if row['Seguimiento']=='Cerrada' else '#ef4444' if row['Seguimiento']=='Pendiente' else '#64748b'
        popup_html = (
            '<div style="font-family:Arial,sans-serif;min-width:175px;line-height:1.65">'
            f'<b>Poste:</b> {poste}<br>'
            f'<b>Maximus:</b> {code}<br>'
            f'<b>Estado:</b> {status_label}'
            '</div>'
        )
        folium.CircleMarker(
            location=[row['Latitud'], row['Longitud']],
            radius=7,
            color='#ffffff',
            weight=2,
            fill=True,
            fill_color=color,
            fill_opacity=1,
            tooltip=f'{poste} · {code}',
            popup=folium.Popup(popup_html, max_width=270),
        ).add_to(group)
    group.add_to(m)
    m.fit_bounds([[points['Latitud'].min(),points['Longitud'].min()],[points['Latitud'].max(),points['Longitud'].max()]], padding=(35,35))
    st_folium(m, height=580, use_container_width=True, returned_objects=[])
    if route:
        st.download_button('Exportar ruta GeoJSON', __import__('json').dumps({'type':'Feature','geometry':{'type':'LineString','coordinates':route['coordinates']},'properties':{'distance_km':route['distance_km']}},ensure_ascii=False), 'ruta_postes.geojson', 'application/geo+json')
