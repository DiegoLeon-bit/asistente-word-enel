# Asistente Enel para Word — versión 0.2

Complemento de prueba (Office Add-in) para Word en la web. Permite completar los cuatro campos básicos y administrar campos adicionales desde el propio panel.

## Qué incluye esta versión

- Campos básicos: `{{NOMBRE_CLIENTE}}`, `{{NUMERO_CASO}}`, `{{SUMINISTRO}}` y `{{RESPUESTA_DE_FONDO}}`.
- Botón **+ Agregar** para crear campos personalizados sin editar JavaScript.
- Botón **Crear campos para etiquetas nuevas**, que aparece después de revisar el documento y permite crear automáticamente un campo por cada etiqueta detectada que aún no esté configurada.
- Configuración de etiquetas personalizadas guardada localmente en el navegador. Los valores que escribas para cada caso no se guardan al cerrar el panel.
- Los campos vacíos se omiten y no borran sus marcadores.

## Actualizar la versión que ya está publicada en GitHub Pages

1. Descarga y descomprime este paquete.
2. Abre tu repositorio `asistente-word-enel` en GitHub.
3. Entra en la carpeta raíz del repositorio (donde están `index.html`, `taskpane.js` y `taskpane.css`).
4. Selecciona **Add file → Upload files**.
5. Sube los archivos actualizados, reemplazando los anteriores: `index.html`, `taskpane.js`, `taskpane.css` y `README.md`. Puedes conservar la carpeta `assets` y el `manifest-template.xml` que ya tienes.
6. Pulsa **Commit changes**.
7. Cuando GitHub Pages termine de publicar, cierra el panel Asistente Enel en Word y vuelve a abrirlo. Si ves la versión anterior, actualiza la página de Word o cierra y vuelve a abrir el documento.

No necesitas volver a cargar el manifiesto en Word si conservas la misma dirección de GitHub Pages.

## Cómo crear campos desde el mismo panel

### Método automático
1. En Word escribe la etiqueta, por ejemplo `{{PERIODO_FACTURADO}}`.
2. En el panel pulsa **Revisar etiquetas**.
3. Pulsa **Crear campos para etiquetas nuevas**.
4. El panel agregará los campos detectados que aún no estén configurados.
5. Escribe el contenido y pulsa **Aplicar al documento**.

### Método manual
1. Pulsa **+ Agregar**.
2. Escribe el nombre del campo, por ejemplo `Periodo facturado`.
3. Escribe la etiqueta exacta, por ejemplo `{{PERIODO_FACTURADO}}`.
4. Escribe el contenido y pulsa **Aplicar al documento**.

Las etiquetas personalizadas deben usar el formato `{{NOMBRE_CAMPO}}`, con letras mayúsculas, números y guion bajo. El texto de la etiqueta debe coincidir exactamente con el que está en Word.

## Privacidad y limitaciones

- El contenido del documento y los valores escritos en el formulario no se envían a un servidor mediante este código.
- Los nombres y las etiquetas de los campos personalizados se guardan en el almacenamiento local del navegador/WebView de Word. Esa configuración no se sincroniza automáticamente con otros computadores o perfiles.
- Los valores de los casos no se guardan en el almacenamiento local.
- La revisión busca marcadores en el texto principal del documento. Encabezados, pies de página, cuadros de texto y otros elementos no están incluidos en esta versión de prueba.
- No necesita Node.js, npm ni un servidor backend.
- Usa Office.js desde el CDN oficial de Microsoft y requiere conexión a Internet para cargar el complemento.

## Publicación inicial y manifiesto

Para la primera publicación, sigue las instrucciones de la versión anterior del proyecto para configurar GitHub Pages y adaptar `manifest-template.xml` a tu usuario de GitHub. Esta versión no cambia la dirección del complemento.
