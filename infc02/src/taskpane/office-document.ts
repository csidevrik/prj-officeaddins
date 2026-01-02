/* global Word console */


export const REQUIRED_TAGS = [
  "ACTA_NUM",
  "PERIODO_NUM",
  "DATE_INIT",
  "DATE_END",
  "CIUDAD",
  "FECHA_ACTA",
  "USD_TOTAL",
  "USD_CONECTIVIDAD",
  "USD_DC",
] as const;

export type RequiredTag = (typeof REQUIRED_TAGS)[number];

export async function validateTemplate(): Promise<Record<string, number>> {
  return Word.run(async (context) => {
    const all = context.document.contentControls;
    all.load("items/tag");
    await context.sync();

    const counts: Record<string, number> = {};
    for (const t of REQUIRED_TAGS) counts[t] = 0;

    for (const cc of all.items) {
      const tag = (cc.tag || "").trim();
      if (!tag) continue;
      counts[tag] = (counts[tag] ?? 0) + 1;
    }

    return counts;
  });
}

export async function applyVars(vars: Record<string, string>): Promise<void> {
  await Word.run(async (context) => {
    for (const [tag, value] of Object.entries(vars)) {
      const ccs = context.document.contentControls.getByTag(tag);
      ccs.load("items");
      await context.sync();

      // si no existe el tag, no revienta; lo reportamos desde la UI con validate
      for (const cc of ccs.items) {
        cc.insertText(value, Word.InsertLocation.replace);
      }
    }
    await context.sync();
  });
}

const insertText = async (text: string) => {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
};

export default insertText;
