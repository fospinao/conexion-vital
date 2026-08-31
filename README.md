# Conexión Vital · Fundación Postobón

App web (sin instalación) tipo búsqueda del tesoro para los parques del
programa. La persona escanea el QR del arco de bienvenida, se registra,
recorre los 4 tótems de la Familia Vital (sin orden fijo), supera los
mini-retos de cada estación y al completar las 4 recibe un código único
para redimir en We Shop.

**URL publicada:** `https://fospinao.github.io/conexion-vital/`

## Estructura

| Archivo | Qué es |
|---|---|
| `index.html` | La app completa (HTML + CSS + JS en un solo archivo) |
| `assets/` | Logos y personajes de la Familia Vital optimizados para web |
| `config/{parque}.json` | Nombre, logo y colores por parque (sin tocar código) |
| `qrs.html` | Generador de los 5 QRs para imprimir (arco + 4 tótems) |
| `backend/apps-script-backend.gs` | Backend completo sobre Google Sheets (instrucciones dentro del archivo) |

## Estaciones

| # | Estación | Anfitrión | Color | Retos |
|---|---|---|---|---|
| 1 | Movimiento | Hija Vital | Verde | Dado + reto físico → Equilibrio → Pregunta express |
| 2 | Conexión | Mamá Vital | Lila | Respiración guiada → Ojo de águila → Video + intención |
| 3 | Bienestar | Papá Vital | Celeste | Clasificar residuos (mín 5/6) → V/F (mín 2/3) |
| 4 | Interacción | Hijo Vital | Naranja | Trivia 15 s (mín 4/5) → Reflejos |

**Contenido variable:** cada participante ve una versión distinta de los
retos (equilibrios, respiraciones, símbolos, residuos, afirmaciones,
preguntas e intenciones se eligen de bancos de opciones usando su
documento + número de recorrido como semilla). Entre personas los retos
cambian, y quien juega otra vez encuentra actividades nuevas (sugerencia
de Colsubsidio). Si alguien reintenta un paso fallido, su reto no cambia.
Los bancos se editan directamente en `index.html` (constantes
`EQUILIBRIOS`, `RESPIRACIONES`, `BANCO_RESIDUOS`, `BANCO_VF`,
`BANCO_TRIVIA`, `INTENCIONES`, etc.).

**Redención (reglas WeShop):** al completar las 4 estaciones se asigna un
código alfanumérico de la base mensual cargada en el Sheet. Máximo **2
códigos por cédula al mes**; cada código tiene **2 meses de vigencia**
(la app muestra la fecha límite). Tras redimir, el botón "Jugar otra
vez" reinicia el pasaporte con retos diferentes.

## Parámetros de URL

- `?parque=comfama|epm|colsubsidio|arvi|...` — branding por parque (default: comfama)
- `&estacion=1..4` — abre directo esa estación (es lo que codifican los QRs)

## Backend (100% Google Sheets)

Un solo Google Sheet es la base de datos, visible y editable por Felipe
y por la Fundación. Pestañas: **parques** (el desplegable del registro se
alimenta de aquí — agregar una fila = parque nuevo en la app, sin tocar
código), **participantes** (una fila por cédula, con progreso, vueltas y
circuitos completos), **estaciones** (log por estación para medir
deserción/embudo), **codigos** (ahí se pega cada mes el Excel de WeShop:
columna `codigo` + columna `mes`) e **intenciones**.

La app funciona en dos modos:

- **Modo local** (CONFIG.SHEET_API vacío): todo vive en `localStorage`
  con las mismas reglas. Sirve para demo y pruebas.
- **Modo producción**: registro, progreso (recuperable desde cualquier
  celular por número de documento) y asignación de códigos pasan por el
  Sheet, con bloqueo para que nunca se repita un código.

### Conectar el backend (una sola vez)

Seguir las instrucciones que están al inicio de
`backend/apps-script-backend.gs` (crear Sheet → pegar script → ejecutar
`instalar` → publicar como Web App → pegar la URL en `CONFIG.SHEET_API`
del `index.html`). Toma unos 5 minutos y no requiere nada técnico
adicional.

### Pasar a producción

En `index.html`: `DEMO_MODE: false` (oculta los botones de simulación de
QR). El video de la estación Conexión se conecta pegando su URL (mp4 o
YouTube embed) en `CONFIG.VIDEO_BIENESTAR`. Los enlaces de la Política de
Datos y los T&C del registro quedan marcados con `href="#"` para conectar
cuando la Fundación entregue los textos definitivos.

## Publicación (GitHub Pages)

Settings → Pages → Source: `main` / (root) → Save.
Para pausar la app: Source → None.

## QRs para los tótems

Abrir `qrs.html` (localmente o publicado), pegar la URL de la app,
elegir el parque y generar. Imprimir con Ctrl+P. El QR es una URL, así
que la cámara nativa del celular siempre funciona como vía alterna al
escáner integrado.
