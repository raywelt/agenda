# Agenda · Semana de la Educación 2026

Página web que muestra la programación de la **XIV Semana Nacional y VI Internacional de la Educación — Educación Expandida** (Colombia, 19 al 24 de octubre de 2026) como calendario, un día a la vez.

- Abre automáticamente el día de hoy o el próximo día con programación. En el día en curso, los eventos que ya terminaron se pliegan y se marcan los que están **en curso** y los **próximos** (hora de Colombia).
- Filtros por **formato, modalidad, sede y lugar**, más un buscador por texto (evento, ponente…).
- Banner horizontal en escritorio y pieza 1x1 en móvil.
- Los filtros y el día quedan en el enlace (`#dia=2026-10-21&sede=calle 80`), así se puede compartir una vista filtrada.

## Cómo actualizar la agenda

La página lee **`data/AgendaPlana.xlsx` directamente**; no hay que convertir nada.

1. Edita el Excel (o reemplázalo por una versión nueva) **con el mismo nombre**: `data/AgendaPlana.xlsx`.
2. Listo. Las páginas abiertas detectan el cambio en menos de un minuto y muestran el aviso "Agenda actualizada"; quien la abra después ya ve la versión nueva.

Reglas para que el Excel se lea bien:

- Puede tener una hoja por día (como ahora) o todo en una sola hoja; se leen todas.
- La fila de encabezados debe tener las columnas `DÍA`, `HORA INICIO`, `HORA FINAL`, `EVENTO`, `LUGAR`, `SEDE`, `FORMATO`, `MODALIDAD` (el orden no importa; mayúsculas y tildes tampoco).
- `DÍA` como fecha de Excel (o texto `19/10/2026`); las horas como hora de Excel (o texto `2:30 pm`). Usa formato de 24 h o indica a. m./p. m.: una hora `3:45` se interpreta como 3:45 de la madrugada.
- En `EVENTO`, la primera línea es el título y las siguientes (Alt+Enter) son los detalles: ponentes, institución, enlaces.
- Las filas "Receso", "Almuerzo" o "Coffee Break" sin formato se muestran como pausas.
- Valores que solo difieren en mayúsculas o tildes ("Por confirmar" / "Por Confirmar") se agrupan en el mismo filtro.

## Ver la página en tu computador

Los navegadores no dejan leer archivos locales desde una página abierta con doble clic, así que hace falta un servidor local mínimo. Desde esta carpeta:

```bash
python3 -m http.server 8000
```

y abre <http://localhost:8000>. (En Windows: `py -m http.server 8000`.)

Si abres `index.html` con doble clic, la página muestra un botón **Cargar AgendaPlana.xlsx** para elegir el archivo a mano.

Para revisar cómo se ve un momento concreto del evento, agrega `?hoy=AAAA-MM-DDTHH:MM` a la dirección, por ejemplo <http://localhost:8000/?hoy=2026-10-20T11:40>.

## Publicarla en internet

Es un sitio estático (HTML, CSS, JS y el Excel), así que sirve en cualquier hosting estático. Con **GitHub Pages**: *Settings → Pages → Deploy from a branch → `main` / root*. Para actualizar la agenda basta con subir el nuevo `AgendaPlana.xlsx` a la carpeta `data/` desde la web de GitHub (*Add file → Upload files*); en un par de minutos queda publicada.

## Archivos

| Archivo | Para qué |
|---|---|
| `index.html` | Estructura de la página |
| `styles.css` | Estilos y colores (tomados de las piezas gráficas) |
| `app.js` | Lectura del Excel, filtros y calendario |
| `data/AgendaPlana.xlsx` | **La agenda — el único archivo que hay que cambiar** |
| `assets/banner.jpg`, `assets/pieza1x1.jpg` | Banner escritorio y pieza móvil |
| `vendor/xlsx.full.min.js` | [SheetJS](https://sheetjs.com) 0.18.5 para leer Excel en el navegador (licencia Apache 2.0) |
