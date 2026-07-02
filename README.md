# Conexión Vital · Postobón — Demo web app

Reto tipo búsqueda del tesoro para parques (Comfama, EPM, Colsubsidio, Arví).
App web sin instalación: registro → 4 estaciones con retos → código de premio.

## Publicar en GitHub Pages (5 minutos)

1. Crea un repositorio nuevo en GitHub (ej: `conexion-vital`), público.
2. Sube TODO el contenido de esta carpeta al repo (index.html, qrs.html, carpeta config/).
3. En el repo: **Settings → Pages → Source: Deploy from a branch → Branch: main / (root) → Save**.
4. En 1–2 minutos la app queda en: `https://TU-USUARIO.github.io/conexion-vital/`

## Probar el flujo completo con cámara

1. Abre `https://TU-USUARIO.github.io/conexion-vital/qrs.html` en el computador.
2. Pega la URL de la app, elige el parque y genera los 5 QRs.
3. Desde el celular: escanea el QR de bienvenida con la cámara nativa → registro.
4. Dentro de la app usa el botón **"Escanear QR"** apuntando a los QRs de estación en la pantalla del PC. La cámara sí funciona aquí porque GitHub Pages es HTTPS.

## Parámetros de URL

- `?parque=comfama|epm|colsubsidio|arvi` → tematización (nombre y colores desde `config/{parque}.json`)
- `&estacion=1..4` → abre directo el reto de esa estación (es lo que llevan los QRs de los tótems)

Agregar un parque nuevo = crear su JSON en `config/` y generar sus QRs. Sin tocar código.

## Pasar a producción (pendientes)

- `CONFIG.DEMO_MODE = false` en index.html (oculta los botones de simulación de QR).
- Conectar backend Firebase en los puntos marcados `>>> INTEGRACIÓN BACKEND <<<`:
  registro, recuperación por documento, estación completada y asignación de código único
  (hoy el progreso y los códigos son locales al dispositivo, solo para demo).
- Cargar la base real de códigos de redención We Shop.
- Reemplazar placeholders de contenido: video "Activa tu bienestar" y audio de meditación.
- Dominio propio si se quiere (ej: retovital.brutalagencia.com) vía CNAME.

---
Brutal Agencia S.A.S. · Demo interna — no distribuir.
