#!/usr/bin/env node
// Build-time segédeszköz: a stdin-en kapott JS-kódot elhomályosítja, és az
// eredményt stdout-ra írja. Kizárólag webapp-client/build.py hívja meg - a
// kimenet (dist_*.html) nem tartalmazza ezt a scriptet, csak a végeredményt.
//
// FIGYELEM: ez - a jelszóhoz és a checksum-ellenőrzéshez hasonlóan - csak
// visszatartás, nem valódi védelem. Nem bontja vissza semmilyen elemzéssel
// szemben, csak nehezebbé/időigényesebbé teszi az olvasást.
import JavaScriptObfuscator from "javascript-obfuscator";

const chunks = [];
for await (const chunk of process.stdin) chunks.push(chunk);
const source = Buffer.concat(chunks).toString("utf-8");

const result = JavaScriptObfuscator.obfuscate(source, {
  compact: true,
  identifierNamesGenerator: "hexadecimal",
  stringArray: true,
  stringArrayEncoding: ["base64"],
  stringArrayThreshold: 0.75,
  simplify: true,
  renameGlobals: false,
  reservedNames: ["^initPage$"],
  controlFlowFlattening: false,
  deadCodeInjection: false,
  selfDefending: false,
  debugProtection: false,
});

process.stdout.write(result.getObfuscatedCode());
