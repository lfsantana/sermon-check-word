# Sermon Check para Word

Complemento de panel lateral para Microsoft Word. Lee el documento abierto, calcula cuánto
dura el sermón o discurso según tu ritmo y coloca etiquetas de tiempo (`[10 MIN]`) dentro
del propio documento.

El texto del documento no sale de Word: todo el cálculo se hace en el panel. Esta página
solo sirve los archivos del panel.

## Instalar (Windows, Word de escritorio)

1. Descarga `manifest.xml` e `Instalar en Word.bat` en la misma carpeta.
2. Haz doble clic en `Instalar en Word.bat`.
3. Cierra Word por completo y vuelve a abrirlo. En la pestaña **Inicio** aparece el botón
   **Sermon Check**.

Para quitarlo, ejecuta `Quitar de Word.bat`.

## Qué hace

- **Duración** estimada y tiempo por punto, a partir de los títulos del documento.
- **Ritmo** configurable: palabras por minuto, ritmo para citas bíblicas, pausa por título
  y margen para improvisar.
- **Etiquetas**: cada X minutos, cada X palabras, al final de cada punto, al 25 %, la mitad
  o el 75 %, o en minutos concretos. También una etiqueta donde está el cursor.
- **Actualizar tiempos**: deja cada etiqueta en su sitio y corrige el minuto tras editar el texto.
- **Modo ensayo**: cronómetro con el texto en grande y aviso de adelanto o atraso.

## Cómo reconoce el documento

- Títulos: párrafos con estilo de título (Título 1, 2, 3…).
- Citas: párrafos con sangría izquierda o estilo «Cita».
- Referencia bíblica: párrafo corto en negrita con capítulo y versículo (`Salmo 19:1`).
- Etiquetas existentes: párrafos que solo contienen `[10 MIN]` o `[10:30]`.
