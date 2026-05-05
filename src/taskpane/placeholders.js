export async function reemplazarPlaceholders(datos) {

  await Word.run(async (context) => {

    // Recorrer todos los placeholders
    for (const key in datos) {

      // Si la key ya tiene el formato {{...}}, la usamos. Si no, le agregamos {{ }}
      const placeholder = key.startsWith("{{") && key.endsWith("}}") ? key : `{{${key}}}`;

      const valor = datos[key];

      // Buscar placeholder
      const resultados = context.document.body.search(
        placeholder,
        {
          matchCase: false,
          matchWholeWord: false
        }
      );

      resultados.load("items");

      await context.sync();

      // Reemplazar cada coincidencia
      resultados.items.forEach((item) => {

        item.insertText(
          valor,
          Word.InsertLocation.replace
        );

      });

    }

    await context.sync();

  });

}