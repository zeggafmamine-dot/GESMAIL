# GESMAIL — RFQ Equipo (demo GitHub Pages)

Demo web del tablero de solicitudes de oferta (RFQ) del departamento comercial.

## Funcionalidades
- Cuatro columnas: **Petición**, **Oferta enviada**, **Oferta perdida** y **Pedido ganado**.
- Importación de correos en formato `.eml`.
- Aviso y bloqueo de duplicados exactos (Message-ID, huella o fichero idéntico).
- Avisos con porcentaje estimado de similitud para correos parecidos.
- Creación manual, edición y cambio de estado de RFQ.
- Búsqueda, exportación y restauración de copias JSON.
- Botón **Cargar ejemplos** con correos ficticios para pruebas.

## Publicar en GitHub Pages
El archivo `index.html` está en la raíz de la rama `main`. Para activar el sitio:
1. Abre **Settings → Pages** en este repositorio.
2. En **Build and deployment**, elige **Deploy from a branch**.
3. Selecciona la rama **main** y la carpeta **/(root)**.
4. Pulsa **Save**.

Dirección prevista: https://zeggafmamine-dot.github.io/GESMAIL/

## Avisos importantes
- **Esta versión NO utiliza Supabase ni una base de datos compartida.**
- Los datos se guardan exclusivamente en el **localStorage de cada navegador**. Los tres compañeros no comparten RFQ en esta demo.
- No hay inicio de sesión ni controles de acceso.
- El repositorio y la página GitHub Pages son **públicos**. **No subas correos ni datos comerciales reales al repositorio.**
- La aplicación procesa los correos localmente, sin enviarlos a un servidor, pero sigue siendo una demostración, no un entorno de producción.
- Solo se admite `.eml` para la importación automática. La compatibilidad con `.msg` y arrastrar desde Outlook directamente están pendientes.
- El porcentaje de parecido es una heurística aproximada, no una garantía de duplicación.

El sitio carga un paquete HTML autocontenido comprimido dentro del propio `index.html`; requiere un navegador moderno con soporte `DecompressionStream`. Su código fuente descomprimido procede del prototipo RFQ Equipo GitHub Pages.