/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

/* otras variables */

// Obtén una referencia al elemento input
const folderInput = document.getElementById("folderInput") as HTMLInputElement;
let dirname = "";
let patdirname = "";


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("findz").onclick = findz;
    // document.getElementById("importd").onclick = importData;
    document.getElementById("printfnameB").onclick = printfname;
     
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

export async function findz() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Buscando Zips, este texto se adiciona al final del texto",Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "green";

    await context.sync();
  });
}


export async function printfname() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph(patdirname,Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}

// Agrega un evento 'change' al elemento input
folderInput.addEventListener("change", () => {
  // Verifica si hay archivos seleccionados
  if (folderInput.files) {
    // Obtén el primer archivo para obtener el directorio
    const firstFile = folderInput.files[0];
    
    // Verifica si es un directorio (webkitRelativePath estará presente en directorios)
    if (firstFile.webkitRelativePath) {
      // Divide el path para obtener el nombre del directorio
      const parts = firstFile.webkitRelativePath.split("/");

      const directoryName = parts[0]; // El nombre del directorio estará en la primera parte
      // Path completo
      patdirname = firstFile.webkitRelativePath; 
      dirname = directoryName;

      // Ahora tienes el nombre del directorio seleccionado
      console.log("Nombre del directorio seleccionado:", directoryName);

      // Puedes hacer lo que desees con el nombre del directorio
    }
  }
});