"use strict";

// Regenerates test/fixtures/smoke-test.pptx: a small but structurally real
// PPTX used for manual/browser smoke testing (drag it into the app) and as
// input for ad-hoc scripts. Not part of `npm test` — the automated tests use
// in-memory JSZip fixtures instead. Run with: node test/fixtures/build-smoke-pptx.js
//
// Contents, by design:
// - 1 slide with a title placeholder (typeface Arial) and a body placeholder
//   (typeface Calibri) -> exercises title/body font replacement
// - slideLayout1 (used) and slideLayout2 (unused) sharing slideMaster1
//   -> exercises unused-layout detection
// - ppt/media/image1.png referenced by a picture on the slide (640x480, displayed ~2")
//   -> exercises display-size ppi image compression
// - ppt/media/image2.png referenced nowhere (orphan, removable)
// - docProps/core.xml + app.xml with a title/creator/company set
//   -> exercises document property scanning/clearing

const path = require("path");
const fs = require("fs");
const zlib = require("zlib");
const JSZip = require("jszip");

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const rel = (id, type, target) => `<Relationship Id="${id}" Type="${REL_NS}/${type}" Target="${target}"/>`;
const relsXml = (rels) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;

// 1x1 transparent PNG
const PNG_1X1 = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
  "base64"
);

function crc32(buf) {
  let crc = ~0 >>> 0;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let bit = 0; bit < 8; bit++) {
      crc = (crc >>> 1) ^ (0xEDB88320 & -(crc & 1));
    }
  }
  return (~crc) >>> 0;
}

function pngChunk(type, data) {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0);
  return Buffer.concat([length, typeBuf, data, crc]);
}

// Oversized raster used on-slide so image compression has something to shrink.
function makeNoisePng(width, height) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  const rows = [];
  for (let y = 0; y < height; y++) {
    const row = Buffer.alloc(1 + width * 3);
    for (let x = 0; x < width; x++) {
      row[1 + x * 3] = Math.floor((x * 255) / Math.max(width - 1, 1));
      row[2 + x * 3] = Math.floor((y * 255) / Math.max(height - 1, 1));
      row[3 + x * 3] = 160;
    }
    rows.push(row);
  }
  const idat = zlib.deflateSync(Buffer.concat(rows));
  const sig = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
  return Buffer.concat([
    sig,
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", idat),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

const PNG_LARGE = makeNoisePng(640, 480);

// Fixed date so re-running this script produces a byte-identical file
// (JSZip stamps each entry with the current time by default, which would
// otherwise make every regeneration show up as a spurious git diff).
const FIXED_DATE = new Date(Date.UTC(2000, 0, 1));

async function build() {
  const zip = new JSZip();
  const addFile = (path, content) => zip.file(path, content, { date: FIXED_DATE });

  addFile(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="png" ContentType="image/png"/>` +
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
    `</Types>`
  );
  addFile("_rels/.rels", relsXml([rel("rId1", "officeDocument", "ppt/presentation.xml")]));

  addFile(
    "ppt/presentation.xml",
    `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<p:sldSz cx="12192000" cy="6858000"/>` +
    `<p:sldMasterIdLst><p:sldMasterId id="1" r:id="rId3"/></p:sldMasterIdLst>` +
    `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
    `</p:presentation>`
  );
  addFile(
    "ppt/_rels/presentation.xml.rels",
    relsXml([
      rel("rId2", "slide", "slides/slide1.xml"),
      rel("rId3", "slideMaster", "slideMasters/slideMaster1.xml"),
    ])
  );

  addFile(
    "ppt/slides/slide1.xml",
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>` +
    `<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:rPr><a:latin typeface="Arial"/></a:rPr><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>` +
    `<p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:rPr><a:latin typeface="Calibri"/></a:rPr><a:t>World</a:t></a:r></a:p></p:txBody></p:sp>` +
    `<p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="457200" y="2743200"/><a:ext cx="1828800" cy="1371600"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>` +
    `</p:spTree></p:cSld></p:sld>`
  );
  addFile(
    "ppt/slides/_rels/slide1.xml.rels",
    relsXml([
      rel("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"),
      rel("rId2", "image", "../media/image1.png"),
    ])
  );

  addFile("ppt/slideLayouts/slideLayout1.xml", "<p:sldLayout/>");
  addFile(
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    relsXml([rel("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")])
  );
  addFile("ppt/slideLayouts/slideLayout2.xml", "<p:sldLayout/>");
  addFile(
    "ppt/slideLayouts/_rels/slideLayout2.xml.rels",
    relsXml([rel("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")])
  );

  addFile("ppt/slideMasters/slideMaster1.xml", "<p:sldMaster/>");

  addFile("ppt/media/image1.png", PNG_LARGE);
  addFile("ppt/media/image2.png", PNG_1X1);

  addFile(
    "docProps/core.xml",
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">` +
    `<dc:title>Smoke Test Deck</dc:title><dc:creator>QA Bot</dc:creator></cp:coreProperties>`
  );
  addFile(
    "docProps/app.xml",
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Company>Acme</Company></Properties>`
  );

  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  const outPath = path.join(__dirname, "smoke-test.pptx");
  fs.writeFileSync(outPath, buffer);
  console.log(`wrote ${buffer.length} bytes to ${outPath}`);
}

build();
