# Asistente Enel para Word — versión inicial 0.1

Complemento de prueba (Office Add-in) para Word en la web. Muestra un panel con cuatro campos y reemplaza en el documento los marcadores correspondientes:

- `{{NOMBRE_CLIENTE}}`
- `{{NUMERO_CASO}}`
- `{{SUMINISTRO}}`
- `{{RESPUESTA_DE_FONDO}}`

## Importante antes de publicarlo

Este paquete contiene código de prueba. No contiene datos reales de clientes. No subas al repositorio documentos, respuestas, nombres, números de suministro ni información interna de Enel.

GitHub Pages publica el sitio en Internet. El código del complemento será visible públicamente. El formulario está diseñado para procesar los valores en el panel y escribirlos en el documento mediante Office.js; esta versión no envía los datos del caso a un servidor.

## Publicar en GitHub Pages (sin Node.js)

1. En GitHub crea un repositorio llamado exactamente `asistente-word-enel`.
2. Para GitHub Free, configura el repositorio como **Public** (GitHub Pages requiere un repositorio público en el plan Free).
3. Descomprime el paquete en tu computador.
4. En el repositorio, selecciona **Add file → Upload files** y carga el contenido de esta carpeta, incluida la carpeta `assets` y todos los archivos. No es necesario cargar el ZIP.
5. Entra a **Settings → Pages**.
6. En **Build and deployment**, selecciona **Deploy from a branch**.
7. Elige la rama `main` y la carpeta `/(root)`, y pulsa **Save**.
8. Espera a que GitHub Pages publique el sitio. La dirección será:
   `https://TU-USUARIO.github.io/asistente-word-enel/`

## Preparar el manifiesto para tu cuenta

1. Abre `manifest-template.xml` con el Bloc de notas o VS Code.
2. Reemplaza todas las apariciones de `TU-USUARIO` por tu nombre de usuario exacto de GitHub.
3. Guarda una copia con el nombre `manifest.xml`.
4. Comprueba en el navegador que se abra la página `https://TU-USUARIO.github.io/asistente-word-enel/`.
5. En Word para la web abre **Inicio → Complementos → Más complementos / Mis complementos → Cargar mi complemento**.
6. Examina y selecciona `manifest.xml`.

Si Word rechaza el manifiesto o muestra un mensaje de política de la organización, no intentes saltarte el control: será necesario solicitar al administrador que permita la carga de complementos personalizados o que lo despliegue mediante el centro de administración de Microsoft 365.

## Probarlo

1. Abre un documento de Word nuevo.
2. Copia el contenido de `plantilla-de-prueba.txt` en el documento. Conserva los marcadores tal como están escritos.
3. Abre el complemento **Asistente Enel para Word** desde Complementos / Mis complementos.
4. Diligencia uno o varios campos y pulsa **Revisar etiquetas**.
5. Pulsa **Aplicar al documento**. Los campos vacíos se omiten; no borran sus marcadores.

## Notas técnicas

- No necesita Node.js, npm ni un servidor backend.
- Usa Office.js desde el CDN oficial de Microsoft.
- Requiere conexión a Internet para cargar la página de GitHub Pages y la biblioteca Office.js.
- Solicita permiso `ReadWriteDocument` porque debe buscar y reemplazar texto en el documento abierto.
- La primera versión busca los marcadores en el texto principal del documento. La revisión visual de encabezados, pies de página, cuadros de texto y otros elementos no está incluida en esta prueba.
- El XML es una plantilla: `TU-USUARIO` debe reemplazarse por el usuario real antes de cargarlo en Word.
