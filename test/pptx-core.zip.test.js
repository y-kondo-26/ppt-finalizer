"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");

const core = require("../docs/pptx-core.js");

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function rel(id, type, target) {
  return `<Relationship Id="${id}" Type="${REL_NS}/${type}" Target="${target}"/>`;
}

function relsXml(rels) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `${rels.join("")}</Relationships>`;
}

// Builds a minimal but structurally realistic PPTX package in memory:
// - slide1 uses slideLayout1 -> slideMaster1, and embeds media/image1.png
// - slideLayout2 is unused by any slide but still embeds media/image3.png
// - slideMaster2 is unused by any layout
// - media/image2.png is referenced nowhere (true package orphan)
// - slide1 has a notes slide, and there is one notes master
async function buildFixtureZip() {
  const zip = new JSZip();

  zip.file(
    "ppt/presentation.xml",
    `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<p:sldMasterIdLst><p:sldMasterId id="1" r:id="rId3"/><p:sldMasterId id="2" r:id="rId5"/></p:sldMasterIdLst>` +
    `<p:notesMasterIdLst><p:notesMasterId r:id="rId6"/></p:notesMasterIdLst>` +
    `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
    `</p:presentation>`
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    relsXml([
      rel("rId2", "slide", "slides/slide1.xml"),
      rel("rId3", "slideMaster", "slideMasters/slideMaster1.xml"),
      rel("rId5", "slideMaster", "slideMasters/slideMaster2.xml"),
      rel("rId6", "notesMaster", "notesMasters/notesMaster1.xml"),
    ])
  );

  zip.file("ppt/slides/slide1.xml", `<p:sld><p:cSld><p:spTree/></p:cSld></p:sld>`);
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    relsXml([
      rel("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"),
      rel("rId2", "image", "../media/image1.png"),
      rel("rId3", "notesSlide", "../notesSlides/notesSlide1.xml"),
    ])
  );

  zip.file("ppt/slideLayouts/slideLayout1.xml", `<p:sldLayout/>`);
  zip.file(
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    relsXml([rel("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")])
  );

  zip.file("ppt/slideLayouts/slideLayout2.xml", `<p:sldLayout/>`);
  zip.file(
    "ppt/slideLayouts/_rels/slideLayout2.xml.rels",
    relsXml([
      rel("rId1", "slideMaster", "../slideMasters/slideMaster1.xml"),
      rel("rId2", "image", "../media/image3.png"),
    ])
  );

  zip.file("ppt/slideMasters/slideMaster1.xml", `<p:sldMaster/>`);
  zip.file("ppt/slideMasters/slideMaster2.xml", `<p:sldMaster/>`);

  zip.file("ppt/media/image1.png", "used-directly");
  zip.file("ppt/media/image2.png", "package-orphan");
  zip.file("ppt/media/image3.png", "used-only-by-unused-layout");

  zip.file("ppt/notesSlides/notesSlide1.xml", `<p:notes/>`);
  zip.file("ppt/notesMasters/notesMaster1.xml", `<p:notesMaster/>`);

  zip.file(
    "docProps/core.xml",
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">` +
    `<dc:title>Test Deck</dc:title><dc:creator>Alice</dc:creator><cp:revision>5</cp:revision>` +
    `</cp:coreProperties>`
  );
  zip.file(
    "docProps/app.xml",
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
    `<Company>Acme</Company></Properties>`
  );

  // Round-trip through generate/load so entries carry real JSZip metadata
  // (compressed/uncompressed sizes), matching how the app loads a real file.
  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  return JSZip.loadAsync(buffer);
}

describe("getPresentationSlideOrder", () => {
  test("maps the single slide to display order 1", async () => {
    const zip = await buildFixtureZip();
    const order = await core.getPresentationSlideOrder(zip);
    assert.equal(order.get("ppt/slides/slide1.xml"), 1);
  });
});

describe("buildDeckStructure / computeCleanupPlan", () => {
  test("flags the unused layout and unused master", async () => {
    const zip = await buildFixtureZip();
    const plan = await core.computeCleanupPlan(zip);

    assert.deepEqual(plan.layoutsToRemove, ["ppt/slideLayouts/slideLayout2.xml"]);
    assert.deepEqual(plan.mastersToRemove, ["ppt/slideMasters/slideMaster2.xml"]);
    assert.equal(plan.unusedLayoutCount, 1);
    assert.equal(plan.unusedMasterCount, 1);
  });

  test("media only reachable through the unused layout is reported as freed by structure cleanup", async () => {
    const zip = await buildFixtureZip();
    const plan = await core.computeCleanupPlan(zip);
    const freedPaths = plan.structureFreedMedia.items.map((item) => item.path);
    assert.deepEqual(freedPaths, ["ppt/media/image3.png"]);
  });

  test("notes and document properties are detected", async () => {
    const zip = await buildFixtureZip();
    const plan = await core.computeCleanupPlan(zip);
    assert.equal(plan.notes.count, 1);
    assert.equal(plan.properties.fieldCount, 4); // title, creator, revision, company
  });
});

describe("computePackageOrphanMedia", () => {
  test("only the never-referenced media file counts as a package orphan", async () => {
    const zip = await buildFixtureZip();
    const orphans = await core.computePackageOrphanMedia(zip);
    const paths = orphans.items.map((item) => item.path);
    assert.deepEqual(paths, ["ppt/media/image2.png"]);
  });
});

describe("analyzeMediaUsage", () => {
  test("labels directly-used media with its slide number", async () => {
    const zip = await buildFixtureZip();
    const slidePathToNum = await core.getPresentationSlideOrder(zip);
    const analysis = await core.analyzeMediaUsage(zip, slidePathToNum);
    const image1 = analysis.items.find((item) => item.path === "ppt/media/image1.png");
    assert.deepEqual(image1.slides, ["スライド 1（直接）"]);
    assert.equal(image1.isOrphan, false);
  });

  test("labels media reachable only through an unused layout as orphaned", async () => {
    const zip = await buildFixtureZip();
    const slidePathToNum = await core.getPresentationSlideOrder(zip);
    const analysis = await core.analyzeMediaUsage(zip, slidePathToNum);
    const image3 = analysis.items.find((item) => item.path === "ppt/media/image3.png");
    assert.equal(image3.isOrphan, true);
    assert.match(image3.slides[0], /^未使用レイアウト/);
  });

  test("media with no relationship at all has no usage labels", async () => {
    const zip = await buildFixtureZip();
    const slidePathToNum = await core.getPresentationSlideOrder(zip);
    const analysis = await core.analyzeMediaUsage(zip, slidePathToNum);
    const image2 = analysis.items.find((item) => item.path === "ppt/media/image2.png");
    assert.deepEqual(image2.slides, []);
    assert.equal(image2.isOrphan, false); // "no usage info" differs from "confirmed orphan"
  });
});

describe("removeNotesFromZip", () => {
  test("deletes the notes slide, notes master, and their relationships", async () => {
    const zip = await buildFixtureZip();
    const result = await core.removeNotesFromZip(zip);

    assert.equal(result.count, 1);
    assert.equal(zip.files["ppt/notesSlides/notesSlide1.xml"], undefined);
    assert.equal(zip.files["ppt/notesMasters/notesMaster1.xml"], undefined);

    const slideRels = await zip.files["ppt/slides/_rels/slide1.xml.rels"].async("string");
    assert.doesNotMatch(slideRels, /notesSlide/);

    const presXml = await zip.files["ppt/presentation.xml"].async("string");
    assert.doesNotMatch(presXml, /notesMasterIdLst/);
  });
});

describe("scanDocumentProperties / clearDocumentProperties", () => {
  test("scan reports the populated fields", async () => {
    const zip = await buildFixtureZip();
    const props = await core.scanDocumentProperties(zip);
    const labels = props.fields.map((f) => f.label);
    assert.ok(labels.includes("タイトル"));
    assert.ok(labels.includes("作成者"));
    assert.ok(labels.includes("会社"));
  });

  test("clear empties the known fields and resets revision to 0", async () => {
    const zip = await buildFixtureZip();
    await core.clearDocumentProperties(zip);

    const coreXml = await zip.files["docProps/core.xml"].async("string");
    assert.equal(core.readXmlElementText(coreXml, "dc:title"), "");
    assert.equal(core.readXmlElementText(coreXml, "cp:revision"), "0");

    const appXml = await zip.files["docProps/app.xml"].async("string");
    assert.equal(core.readXmlElementText(appXml, "Company"), "");
  });
});

describe("removeUnusedStructure", () => {
  test("removes the unused layout/master parts and detaches them from their owners", async () => {
    const zip = await buildFixtureZip();
    const plan = await core.computeCleanupPlan(zip);
    const result = await core.removeUnusedStructure(zip, plan);

    assert.equal(result.layoutsRemoved, 1);
    assert.equal(result.mastersRemoved, 1);
    assert.equal(zip.files["ppt/slideLayouts/slideLayout2.xml"], undefined);
    assert.equal(zip.files["ppt/slideMasters/slideMaster2.xml"], undefined);

    const presRels = await zip.files["ppt/_rels/presentation.xml.rels"].async("string");
    assert.doesNotMatch(presRels, /slideMaster2\.xml/);
  });

  test("media used only by the removed unused layout is deleted once the follow-up orphan sweep runs", async () => {
    // Mirrors the real app flow: finalizePptx() calls removeUnusedStructure()
    // and then removePackageOrphanMedia() right after, so a layout's media
    // only becomes "orphan" (and gets swept) once the layout's own .rels
    // file is gone. image3.png in the fixture is referenced only by the
    // unused slideLayout2.
    const zip = await buildFixtureZip();
    const plan = await core.computeCleanupPlan(zip);

    await core.removeUnusedStructure(zip, plan);
    const sweep = await core.removePackageOrphanMedia(zip);

    assert.equal(zip.files["ppt/media/image3.png"], undefined);
    assert.ok(sweep.count >= 1);
    // media still used directly by a live slide must survive the sweep
    assert.notEqual(zip.files["ppt/media/image1.png"], undefined);
  });
});

describe("computeReductionEstimate", () => {
  test("counts embedded font bytes only when replaceFonts is enabled", async () => {
    const zip = new JSZip();
    zip.file("ppt/fonts/font1.fntdata", "0123456789");
    const buffer = await zip.generateAsync({ type: "nodebuffer" });
    const loaded = await JSZip.loadAsync(buffer);

    const plan = {
      slideOrphanMedia: { items: [] },
      layoutsToRemove: [],
      mastersToRemove: [],
      structureFreedMedia: { totalCompressedSize: 0 },
      notes: { count: 0, bytes: 0, compressedBytes: 0 },
      properties: { removableBytes: 0 },
    };
    const baseOptions = {
      removeOrphanMedia: false,
      removeUnusedStructure: false,
      removeNotes: false,
      removeProperties: false,
    };

    const withFonts = core.computeReductionEstimate(plan, { ...baseOptions, replaceFonts: true }, loaded);
    const withoutFonts = core.computeReductionEstimate(plan, { ...baseOptions, replaceFonts: false }, loaded);

    assert.ok(withFonts > 0);
    assert.equal(withoutFonts, 0);
  });

  test("adds image-compress savings when compressImages is enabled", () => {
    const zip = new JSZip();
    const plan = {
      slideOrphanMedia: { items: [] },
      layoutsToRemove: [],
      mastersToRemove: [],
      structureFreedMedia: { totalCompressedSize: 0 },
      notes: { count: 0, bytes: 0, compressedBytes: 0 },
      properties: { removableBytes: 0 },
      imageCompressUsages: [{
        path: "ppt/media/photo.jpg",
        width: 4000,
        height: 4000,
        mime: "image/jpeg",
        size: 500000,
        compressedSize: 400000,
        uses: [{
          displayCxEmu: core.EMU_PER_INCH,
          displayCyEmu: core.EMU_PER_INCH,
          visibleRatioW: 1,
          visibleRatioH: 1,
        }],
      }],
    };
    const base = {
      replaceFonts: false,
      removeOrphanMedia: false,
      removeUnusedStructure: false,
      removeNotes: false,
      removeProperties: false,
    };
    const off = core.computeReductionEstimate(plan, { ...base, compressImages: false }, zip);
    const on = core.computeReductionEstimate(plan, {
      ...base,
      compressImages: true,
      imagePpi: 150,
      jpegQuality: 0.75,
    }, zip);
    assert.equal(off, 0);
    assert.ok(on > 0);
  });
});

describe("collectImageCompressJobs / compressImagesInZip", () => {
  function pngStub(width, height, fill = 0) {
    const buf = Buffer.alloc(64 + fill);
    buf[0] = 0x89;
    buf.write("PNG\r\n\x1a\n", 1);
    buf.writeUInt32BE(13, 8);
    buf.write("IHDR", 12);
    buf.writeUInt32BE(width, 16);
    buf.writeUInt32BE(height, 20);
    buf.fill(0x61, 24);
    return buf;
  }

  function cropXml(srcRect) {
    return srcRect
      ? `<a:srcRect l="${srcRect.l || 0}" t="${srcRect.t || 0}" r="${srcRect.r || 0}" b="${srcRect.b || 0}"/>`
      : "";
  }

  function picXml(rId, cx, cy, srcRect) {
    const crop = cropXml(srcRect);
    return `<p:pic><p:blipFill><a:blip r:embed="${rId}"/>${crop}</p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr></p:pic>`;
  }

  function shapeFillXml(rId, cx, cy, srcRect) {
    return `<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:blipFill><a:blip r:embed="${rId}"/>${cropXml(srcRect)}</a:blipFill></p:spPr></p:sp>`;
  }

  function tableFillXml(rId, cx, cy, srcRect) {
    return `<p:graphicFrame><p:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>` +
      `<a:graphic><a:graphicData><a:tbl>` +
      `<a:tblGrid><a:gridCol w="${cx}"/></a:tblGrid>` +
      `<a:tr h="${cy}"><a:tc><a:tcPr>` +
      `<a:blipFill><a:blip r:embed="${rId}"/>${cropXml(srcRect)}</a:blipFill>` +
      `</a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
  }

  async function buildImageZip({ photos }) {
    const zip = new JSZip();
    zip.file(
      "ppt/presentation.xml",
      `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<p:sldSz cx="12192000" cy="6858000"/>` +
      `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>`
    );
    zip.file(
      "ppt/_rels/presentation.xml.rels",
      relsXml([rel("rId2", "slide", "slides/slide1.xml")])
    );

    const pics = [];
    const imageRels = [];
    photos.forEach((photo, index) => {
      const rId = `rId${index + 2}`;
      const mediaName = photo.name;
      const drawing = photo.kind === "shape"
        ? shapeFillXml(rId, photo.cx, photo.cy, photo.srcRect)
        : photo.kind === "table"
          ? tableFillXml(rId, photo.cx, photo.cy, photo.srcRect)
          : picXml(rId, photo.cx, photo.cy, photo.srcRect);
      pics.push(photo.wrap ? photo.wrap(drawing) : drawing);
      imageRels.push(rel(rId, "image", `../media/${mediaName}`));
      zip.file(`ppt/media/${mediaName}`, photo.bytes);
    });

    zip.file(
      "ppt/slides/slide1.xml",
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<p:cSld><p:spTree>${pics.join("")}</p:spTree></p:cSld></p:sld>`
    );
    zip.file(
      "ppt/slides/_rels/slide1.xml.rels",
      relsXml([
        rel("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"),
        ...imageRels,
      ])
    );
    const buffer = await zip.generateAsync({ type: "nodebuffer" });
    return JSZip.loadAsync(buffer);
  }

  test("assigns a smaller target to an icon than to a full-slide photo", async () => {
    const zip = await buildImageZip({
      photos: [
        {
          name: "hero.png",
          bytes: pngStub(4000, 2250, 200),
          cx: 12192000,
          cy: 6858000,
        },
        {
          name: "icon.png",
          bytes: pngStub(4000, 4000, 200),
          cx: core.EMU_PER_INCH,
          cy: core.EMU_PER_INCH,
        },
      ],
    });
    const jobs = await core.collectImageCompressJobs(zip, { imagePpi: 150 });
    const hero = jobs.find((job) => job.path.endsWith("hero.png"));
    const icon = jobs.find((job) => job.path.endsWith("icon.png"));
    assert.ok(hero.targetWidth > icon.targetWidth);
    assert.equal(icon.targetWidth, 150);
  });

  test("shared media uses the largest display size", async () => {
    const zip = await buildImageZip({
      photos: [
        {
          name: "shared.png",
          bytes: pngStub(2000, 2000, 80),
          cx: core.EMU_PER_INCH,
          cy: core.EMU_PER_INCH,
        },
        {
          name: "shared.png",
          bytes: pngStub(2000, 2000, 80),
          cx: core.EMU_PER_INCH * 4,
          cy: core.EMU_PER_INCH * 4,
        },
      ],
    });
    const jobs = await core.collectImageCompressJobs(zip, { imagePpi: 150 });
    const shared = jobs.filter((job) => job.path.endsWith("shared.png"));
    assert.equal(shared.length, 1);
    assert.equal(shared[0].targetWidth, 600);
  });

  test("srcRect increases needed pixels for the stored image", async () => {
    const zip = await buildImageZip({
      photos: [
        {
          name: "cropped.png",
          bytes: pngStub(2000, 1000, 80),
          cx: core.EMU_PER_INCH,
          cy: core.EMU_PER_INCH / 2,
          srcRect: { l: 25000, r: 25000, t: 0, b: 0 },
        },
      ],
    });
    const [job] = await core.collectImageCompressJobs(zip, { imagePpi: 150 });
    assert.equal(job.targetWidth, 300);
  });

  test("replaces media only when the codec returns a smaller payload", async () => {
    const zip = await buildImageZip({
      photos: [{
        name: "hero.png",
        bytes: pngStub(800, 800, 400),
        cx: core.EMU_PER_INCH,
        cy: core.EMU_PER_INCH,
      }],
    });
    const before = await zip.files["ppt/media/hero.png"].async("uint8array");

    await core.compressImagesInZip(zip, { compressImages: true, imagePpi: 150 }, {
      async encode() {
        return { bytes: new Uint8Array(before.byteLength + 10) };
      },
    });
    const unchanged = await zip.files["ppt/media/hero.png"].async("uint8array");
    assert.equal(unchanged.byteLength, before.byteLength);

    const smaller = new Uint8Array(10).fill(7);
    const result = await core.compressImagesInZip(zip, { compressImages: true, imagePpi: 150 }, {
      async encode() {
        return { bytes: smaller };
      },
    });
    assert.equal(result.count, 1);
    const after = await zip.files["ppt/media/hero.png"].async("uint8array");
    assert.deepEqual(Buffer.from(after), Buffer.from(smaller));
  });

  test("collects a shape picture fill as a compress job", async () => {
    const zip = await buildImageZip({
      photos: [{
        name: "fill.png",
        kind: "shape",
        bytes: pngStub(4000, 4000, 80),
        cx: core.EMU_PER_INCH,
        cy: core.EMU_PER_INCH,
      }],
    });
    const [job] = await core.collectImageCompressJobs(zip, { imagePpi: 150 });
    assert.equal(job.path, "ppt/media/fill.png");
    assert.equal(job.targetWidth, 150);
  });

  test("collects a table cell picture fill as a compress job", async () => {
    const zip = await buildImageZip({
      photos: [{
        name: "cell.png",
        kind: "table",
        bytes: pngStub(4000, 2000, 80),
        cx: core.EMU_PER_INCH,
        cy: core.EMU_PER_INCH / 2,
      }],
    });
    const [job] = await core.collectImageCompressJobs(zip, { imagePpi: 150 });
    assert.equal(job.path, "ppt/media/cell.png");
    assert.equal(job.targetWidth, 150);
    assert.equal(job.targetHeight, 75);
  });

  test("walks pictures inside a group even when grpSpPr xfrm is missing", async () => {
    const zip = await buildImageZip({
      photos: [{
        name: "grouped.png",
        bytes: pngStub(4000, 4000, 80),
        cx: core.EMU_PER_INCH,
        cy: core.EMU_PER_INCH,
        wrap: (inner) => `<p:grpSp><p:grpSpPr></p:grpSpPr>${inner}</p:grpSp>`,
      }],
    });
    const [job] = await core.collectImageCompressJobs(zip, { imagePpi: 150 });
    assert.equal(job.path, "ppt/media/grouped.png");
    assert.equal(job.targetWidth, 150);
  });
});

describe("readZipEntryPrefix", () => {
  test("returns only the requested prefix of a stored entry", async () => {
    const zip = new JSZip();
    const payload = Buffer.alloc(80 * 1024, 7);
    payload[0] = 0x89;
    payload.write("PNG\r\n\x1a\n", 1);
    payload.writeUInt32BE(13, 8);
    payload.write("IHDR", 12);
    payload.writeUInt32BE(640, 16);
    payload.writeUInt32BE(480, 20);
    zip.file("ppt/media/big.png", payload, { compression: "STORE" });
    const loaded = await JSZip.loadAsync(await zip.generateAsync({
      type: "nodebuffer",
      compression: "STORE",
    }));

    const prefix = await core.readZipEntryPrefix(loaded.files["ppt/media/big.png"], 64);
    assert.equal(prefix.length, 64);
    assert.equal(prefix[0], 0x89);
    const dims = core.readImageDimensions(prefix);
    assert.deepEqual(dims, { width: 640, height: 480 });
  });
});
