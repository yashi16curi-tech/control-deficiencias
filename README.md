# Dashboard minimalista de seguimiento de deficiencias

## Uso en Windows / Visual Studio Code

1. Descomprime el proyecto.
2. En terminal, entra a la carpeta donde está `app.py` y ejecuta:

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

3. Coloca la carpeta `INSPECCIONES` al mismo nivel que `app.py`, o indica su ruta completa en el panel izquierdo.

Ejemplo de estructura:

```text
DASHBOARD/
  dashboard_minimalista/
    app.py
    requirements.txt
  INSPECCIONES/
    770125451012/
      antes.jpg
      despues.jpg
    770125461002/
      01.jpg
      02.jpg
```

También se admite `INSPECCIONES.zip`, con subcarpetas nombradas por código Maximus.

## Cómo funcionan los filtros

Los filtros son dependientes en este orden:

`Unidad → Alimentador → Tipificación → Seguimiento`

Cada selector muestra solo opciones válidas según la selección anterior.

## Criterio de cierre

El **estado documental** se calcula con el **Word**:

- **Cerrada**: la ficha del código Maximus tiene las 2 fotos (antes y después).
- **Pendiente**: la ficha existe, pero falta una de las fotos.
- **Sin evidencia**: el código no fue encontrado en el Word.

La carpeta `INSPECCIONES` se usa **solo para visualizar imágenes** y ya **no cambia** el estado del registro.

## Cambios de la versión actual

- Se eliminó la tabla grande principal del dashboard anterior y se dejó una vista más compacta.
- La sección **Por alimentador** ahora es gráfica.
- Los botones `Ver cerradas`, `Ver pendientes` y `Sin evidencia` filtran directamente la vista.
- El conteo de pendientes ya no se infla ni se reduce por la cantidad de fotos en la carpeta `INSPECCIONES`.

## Cómo actualizar la información

### Nuevo Excel
Cárgalo en el panel izquierdo en **Excel actualizado**.

### Nuevo Word
Cárgalo en el panel izquierdo en **Reporte Word actualizado**.

### Nuevas fotos
- Si usas carpeta local: agrega fotos a `INSPECCIONES` y recarga la app.
- Si usas ZIP: vuelve a cargar `INSPECCIONES.zip`.

Si reemplazas el Excel o el Word por archivos más recientes, el dashboard recalcula automáticamente los indicadores y gráficos.

## Vista de ficha de intervención (v4)

La sección de detalle reproduce la organización del Word: cuadro de identificación a la izquierda,
actividad/OM/responsable a la derecha y abajo dos paneles **Antes** y **Después** del mismo tamaño.
Si la carpeta del código Maximus contiene solo una imagen, el otro panel aparece vacío.
Se buscan nombres `antes` y `despues` cuando existen; sin esos nombres se toma el orden
alfabético de las imágenes de la carpeta. Para evitar confusiones, se recomienda renombrarlas
`antes.jpg` y `despues.jpg`.

Los estados de cumplimiento siguen viniendo del Word; estas fotos de carpeta solo sirven
para la visualización de la ficha.

## Regla de fotos — versión 5

- `Foto_antes` se muestra en el recuadro **Foto antes**.
- `Detalle_antes` se integra como acercamiento circular sobre la fotografía **Antes**, tal como en Word.
- Solo un archivo con `despues`, `posterior` o `final` en su nombre se asigna al recuadro **Foto después**.
- Si solo existen `Foto_antes` y `Detalle_antes`, el recuadro posterior se mantiene vacío.
- La clasificación de cierre sigue tomando como referencia las evidencias del **Word**. Las fotos de `INSPECCIONES` no alteran los indicadores.
