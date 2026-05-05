async function insertarArchivo(nombreArchivo, ubicacion) {
  try {
    const response = await fetch(`/assets/${nombreArchivo}`);
    const arrayBuffer = await response.arrayBuffer();
    const base64String = arrayBufferToBase64(arrayBuffer);

    await Word.run(async (context) => {
      if (ubicacion === "header") {
        const header = context.document.sections.getFirst().getHeader("Primary");
        header.insertFileFromBase64(base64String, Word.InsertLocation.replace);
      } else if (ubicacion === "start") {
        context.document.body.insertFileFromBase64(base64String, Word.InsertLocation.start);
      } else {
        // Usa Word.InsertLocation.end para añadir al final, o .replace si deseas sobreescribir
        context.document.body.insertFileFromBase64(base64String, Word.InsertLocation.end);
      }
      await context.sync();
    });

    console.log(`${nombreArchivo} insertado correctamente`);
  } catch (error) {
    console.error(`Error insertando ${nombreArchivo}:`, error);
  }
}

export async function insertarEncabezado() {
  // Inserta el archivo al principio de la hoja
  await insertarArchivo("Encabezado.docx", "start");
}

export async function insertarCuerpo() {
  // Inserta el archivo en el cuerpo del documento
  await insertarArchivo("Cuerpo.docx", "body");
}


// Convertir ArrayBuffer → Base64
function arrayBufferToBase64(buffer) {

  let binary = "";

  const bytes = new Uint8Array(buffer);

  const length = bytes.byteLength;

  for (let i = 0; i < length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }

  return window.btoa(binary);

}