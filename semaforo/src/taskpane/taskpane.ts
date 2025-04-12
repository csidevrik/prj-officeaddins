/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");

    if (sideloadMsg && appBody) {
      sideloadMsg.style.display = "none";
      appBody.style.display = "flex";
    }

    const runBtn = document.getElementById("run");
    if (runBtn) runBtn.onclick = run;
  }
});


export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}

export async function changeColor(color: string) {
  const colors: Record<string, string> = {
    rojo: "#FF0000",
    ambar: "#FFC000",
    ambar_critico: "#FF6600",
    verde: "#00B050",
  };
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.font.color = colors[color];
      await context.sync();
    });
  } catch (error) {
    console.error("Error changing color:", error);
  }
}
// (window as any).changeColor = changeColor;

export async function listarShapesHeader(){
  try {
    console.log("🚀 Ejecutando listarShapesHeader...");
    await Word.run(async (context) => {
      const secciones = context.document.sections;
      context.load(secciones, "items");
      await context.sync();
      if (secciones.items.length === 0) {
        console.log("No hay secciones en el documento.");
        return;
      }

      const header = secciones.items[0].getHeader(Word.HeaderFooterType.primary);
      const shapes = header.shapes;
      context.load(shapes,"items/name,items/textFrame/textRange/text");
      // context.load(shapes, "items");

      await context.sync();

      if (shapes.items.length === 0) {
        console.warn("⚠️ No hay formas en el encabezado.");
      } else {
        console.log(`🔎 Se encontraron ${shapes.items.length} formas en el encabezado:`);
        // Cargar propiedades individuales por separado
        for (let i = 0; i < shapes.items.length; i++) {
          const shape = shapes.items[i];
          context.load(shape, "name, textFrame/textRange/text");
        }
        await context.sync();
        shapes.items.forEach((shape,index) => {
          const texto = shape.textFrame?.textRange?.text || "Sin texto";
          console.log(`→ Shape ${index + 1}`);
          console.log(`   - Nombre: ${shape.name}`);
          console.log(`   - Texto: ${texto}`);
      });
      }
    });
  } catch (error) {
    console.error("Error listing shapes:", error);
    OfficeHelpers.UI.notifyUser("An error occurred while listing shapes. Please try again.");
  }
  
}

export async function listarShapesEnTodasLasCabeceras() {
  try {
    console.log("🚀 Ejecutando listarShapesEnTodasLasCabeceras...");
    await Word.run(async (context) => {
      const secciones = context.document.sections;
      context.load(secciones, "items");
      await context.sync();

      if (secciones.items.length === 0) {
        console.log("No hay secciones en el documento.");
        return;
      }

      for (let i = 0; i < secciones.items.length; i++) {
        const seccion = secciones.items[i];
        console.log(`🔎 Procesando sección ${i + 1}...`);

        const header = seccion.getHeader(Word.HeaderFooterType.primary);
        const shapes = header.shapes;
        context.load(shapes, "items/name,items/textFrame/textRange/text");

        await context.sync();

        if (shapes.items.length === 0) {
          console.warn(`⚠️ No hay formas en el encabezado de la sección ${i + 1}.`);
        } else {
          console.log(`✅ Se encontraron ${shapes.items.length} formas en el encabezado de la sección ${i + 1}:`);
          for (let j = 0; j < shapes.items.length; j++) {
            const shape = shapes.items[j];
            context.load(shape, "name, textFrame/textRange/text");
          }
          await context.sync();

          shapes.items.forEach((shape, index) => {
            const texto = shape.textFrame?.textRange?.text || "Sin texto";
            console.log(`→ Shape ${index + 1} en sección ${i + 1}`);
            console.log(`   - Nombre: ${shape.name}`);
            console.log(`   - Texto: ${texto}`);
          });
        }
      }
    });
  } catch (error) {
    console.error("Error listing shapes in all headers:", error);
    console.error("An error occurred while listing shapes in all headers. Please try again.");
  }
}

export async function listarTodosLosHeadersShapes() {
  try {
    await Word.run(async (context) => {
      const secciones = context.document.sections;
      context.load(secciones, "items");

      await context.sync();

      const headersToCheck = [
        { type: Word.HeaderFooterType.primary, name: "Primary" },
        { type: Word.HeaderFooterType.firstPage, name: "First Page" },
        { type: Word.HeaderFooterType.evenPages, name: "Even Pages" },
      ];

      for (const { type, name } of headersToCheck) {
        const header = secciones.items[0].getHeader(type);
        const shapes = header.shapes;
        context.load(shapes, "items");

        await context.sync();

        console.log(`🧾 ${name} header: ${shapes.items.length} shape(s)`);
        for (let i = 0; i < shapes.items.length; i++) {
          const shape = shapes.items[i];
          context.load(shape, "name, textFrame/textRange/text");
        }

        await context.sync();

        shapes.items.forEach((shape, index) => {
          const texto = shape.textFrame?.textRange?.text || "Sin texto";
          console.log(`→ [${name}] Shape ${index + 1}`);
          console.log(`   - Nombre: ${shape.name}`);
          console.log(`   - Texto: ${texto}`);
        });
      }
    });
  } catch (error) {
    console.error("❌ Error en listarTodosLosHeadersShapes:", error);
  }
}

(window as any).listarTodosLosHeadersShapes = listarTodosLosHeadersShapes;



declare global {
  interface Window {
    changeColor: (color: string) => Promise<void>;
    listarShapesHeader: () => Promise<void>;
  }
}
window.changeColor = changeColor;
window.listarShapesHeader = listarShapesHeader;
// (window as any).listarShapesHeader = listarShapesHeader;


