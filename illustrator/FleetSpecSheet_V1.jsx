/*
Fleet Spec Sheet Generator — Illustrator (V1)
- Builds spec sheet artboard(s) INSIDE the active document.
- Sources PROOF + PRINT items from ARTBOARDS (no file paths).
- Prompts only for: Customer + Vehicle (single dialog, 2 fields).
- Vehicle Unit # always blank for handwriting.
- HudsonNY font everywhere (falls back if missing).

Artboard rules:
- Proof artboard name: "PROOF" or "MOCKUP" (case-insensitive) = proof image
- Ignored artboards (case-insensitive): names matching "unit number", "unit numbers", "unit #", "unit no"
- All other artboards become Print Manifest rows
- Qty parsing from artboard name: QTY2, QTY_2, QTY-2, QTY 2, x2, 2x
*/

(function () {
  if (app.documents.length === 0) {
    alert("Open a document first.");
    return;
  }

  var doc = app.activeDocument;

  // --- constants (points) ---
  var IN = 72;
  var PAGE_W = 8.5 * IN;
  var PAGE_H = 11 * IN;
  var MARGIN = 0.5 * IN;
  var CONTENT_W = PAGE_W - 2 * MARGIN; // 7.5"
  var X0 = MARGIN;
  var TOP = PAGE_H - MARGIN;
  var BOT = MARGIN;

  // header
  var HEADER_H = 0.45 * IN; // ~32pt
  var HEADER_TEXT = "BUFFALO GRAPHICS COMPANY";

  // details table height
  var DETAILS_H = 0.6 * IN; // ~43pt

  // proof box
  var PROOF_OUTER_W = CONTENT_W; // 7.5"
  var PROOF_INNER_W = 7.2 * IN; // 7.2"
  var PROOF_GAP_TOP = 0.14 * IN;

  // manifest
  var MANIFEST_TITLE_GAP = 0.12 * IN;
  var MANIFEST_HDR_H = 0.32 * IN;
  var ROW_H = 0.85 * IN;
  var COL_W = {
    thumb: 1.4 * IN,
    name: 1.9 * IN,
    file: 2.6 * IN,
    qty: 0.7 * IN,
    done: 0.9 * IN,
  };
  var THUMB_BOX_W = 0.95 * IN;
  var THUMB_BOX_H = 0.7 * IN;

  // sign-off
  var SIGN_TITLE_GAP = 0.2 * IN;
  var SIGN_H = 1.55 * IN;

  // Active spec-sheet artboard origin (set per page)
  var __AB_LEFT = 0;
  var __AB_TOP = 0;

  function toDocX(x) {
    return __AB_LEFT + x;
  }

  function toDocY(y) {
    // Convert page Y (0..PAGE_H from bottom) into Illustrator document Y
    return __AB_TOP - (PAGE_H - y);
  }

  // --- helpers ---
  function findFontContains(substr) {
    try {
      var needle = String(substr || "").toLowerCase();
      if (!needle) return null;
      for (var i = 0; i < app.textFonts.length; i++) {
        var f = app.textFonts[i];
        var n = f && f.name ? String(f.name).toLowerCase() : "";
        var fam = f && f.family ? String(f.family).toLowerCase() : "";
        var sty = f && f.style ? String(f.style).toLowerCase() : "";
        if (
          n.indexOf(needle) !== -1 ||
          fam.indexOf(needle) !== -1 ||
          sty.indexOf(needle) !== -1
        )
          return f;
      }
    } catch (e) {}
    return null;
  }
  function tryFont(name) {
    try {
      return app.textFonts.getByName(name);
    } catch (e) {
      return null;
    }
  }

  var HUDSON =
    tryFont("HudsonNY") ||
    tryFont("Hudson NY") ||
    tryFont("HudsonNY-Regular") ||
    findFontContains("hudson");

  if (!HUDSON) {
    alert(
      "HudsonNY font not found in Illustrator. Install/activate the font, then restart Illustrator.\n\n" +
        "Using default font for now."
    );
  }

  function setText(tf, size, colorRGB, fontObj) {
    var tr = tf.textRange;
    tr.characterAttributes.size = size;
    if (fontObj) tr.characterAttributes.textFont = fontObj;
    else if (HUDSON) tr.characterAttributes.textFont = HUDSON;
    if (colorRGB) {
      var c = new RGBColor();
      c.red = colorRGB[0];
      c.green = colorRGB[1];
      c.blue = colorRGB[2];
      tr.characterAttributes.fillColor = c;
    }
  }

  function strokeBlack(item, w) {
    item.stroked = true;
    item.strokeWidth = w || 1;
    var c = new RGBColor();
    c.red = 0;
    c.green = 0;
    c.blue = 0;
    item.strokeColor = c;
    item.filled = false;
  }

  function fillBlack(item) {
    item.filled = true;
    var c = new RGBColor();
    c.red = 0;
    c.green = 0;
    c.blue = 0;
    item.fillColor = c;
    item.stroked = false;
  }

  function addRect(layer, x, yTop, w, h, fill) {
    // Illustrator rect uses (left, top, width, height)
    var r = layer.pathItems.rectangle(yTop, x, w, h);
    if (fill) fillBlack(r);
    else strokeBlack(r, 1);
    return r;
  }

  function addLine(layer, x1, y1, x2, y2) {
    var ln = layer.pathItems.add();
    ln.setEntirePath([
      [x1, y1],
      [x2, y2],
    ]);
    strokeBlack(ln, 1);
    return ln;
  }

  function addText(layer, x, yTop, w, h, s, size, rgb, fontObj, align) {
    // Use AREA TEXT so width/height are valid (point text cannot set width/height)
    var box = layer.pathItems.rectangle(yTop, x, w, h);
    box.stroked = false;
    box.filled = false;

    var tf = layer.textFrames.areaText(box);
    tf.contents = s;

    setText(tf, size, rgb, fontObj);

    if (align) {
      try {
        tf.textRange.paragraphAttributes.justification = align;
      } catch (e) {}
    }

    return tf;
  }

  function parseQty(name) {
    var n = name;
    var m;
    m = n.match(/(?:^|[_\-\s])qty(?:[_\-\s]*)(\d+)(?:[_\-\s\.]|$)/i);
    if (m) return parseInt(m[1], 10);
    m = n.match(/(?:^|[_\-\s])x(\d+)(?:[_\-\s\.]|$)/i);
    if (m) return parseInt(m[1], 10);
    m = n.match(/(?:^|[_\-\s])(\d+)x(?:[_\-\s\.]|$)/i);
    if (m) return parseInt(m[1], 10);
    return 1;
  }

  function isIgnoredArtboard(name) {
    return /unit\s*(number|numbers|#|no\.?)/i.test(name);
  }

  function isProofArtboard(name) {
    return (
      /^\s*(proof|mockup)\s*$/i.test(name) ||
      /(^|\s)(proof|mockup)($|\s)/i.test(name)
    );
  }

  function exportArtboardPNG(doc, abIndex, outFile) {
    var prev = doc.artboards.getActiveArtboardIndex();
    doc.artboards.setActiveArtboardIndex(abIndex);

    var opt = new ExportOptionsPNG24();
    opt.antiAliasing = true;
    opt.transparency = true;
    opt.artBoardClipping = true;
    opt.horizontalScale = 200; // percent
    opt.verticalScale = 200;

    doc.exportFile(outFile, ExportType.PNG24, opt);
    doc.artboards.setActiveArtboardIndex(prev);
  }

  function fitPlacedToBox(placed, boxW, boxH) {
    // scale to fit (keep aspect)
    var vb = placed.visibleBounds; // [left, top, right, bottom]
    var w = vb[2] - vb[0];
    var h = vb[1] - vb[3];
    if (w <= 0 || h <= 0) return;

    var s = Math.min(boxW / w, boxH / h) * 100;
    placed.resize(s, s); // percent
  }

  function centerPlacedInCell(placed, cellX, cellYTop, cellW, cellH) {
    var vb = placed.visibleBounds;
    var w = vb[2] - vb[0];
    var h = vb[1] - vb[3];

    var cx = cellX + cellW / 2;
    var cy = cellYTop - cellH / 2;

    placed.left = toDocX(cx - w / 2);
    placed.top = toDocY(cy + h / 2);
  }

  function makeDetailsDialog(defaultCustomer, defaultVehicle) {
    var w = new Window("dialog", "Spec Sheet Details");
    w.orientation = "column";
    w.alignChildren = ["fill", "top"];

    var g1 = w.add("group");
    g1.add("statictext", undefined, "Customer:");
    var customer = g1.add("edittext", undefined, defaultCustomer || "");
    customer.characters = 38;

    var g2 = w.add("group");
    g2.add("statictext", undefined, "Vehicle:");
    var vehicle = g2.add("edittext", undefined, defaultVehicle || "MIXER");
    vehicle.characters = 38;

    var btns = w.add("group");
    btns.alignment = "right";
    btns.add("button", undefined, "Cancel", { name: "cancel" });
    var ok = btns.add("button", undefined, "OK", { name: "ok" });

    customer.active = true;
    if (w.show() !== 1) return null;

    return {
      customer: (customer.text || "").replace(/^\s+|\s+$/g, ""),
      vehicle: (vehicle.text || "").replace(/^\s+|\s+$/g, "") || "MIXER",
    };
  }

  // --- gather artboards ---
  var proofIdx = -1;
  var items = []; // {name, qty, abIndex}

  for (var i = 0; i < doc.artboards.length; i++) {
    var abName = doc.artboards[i].name || "Artboard " + (i + 1);
    if (isIgnoredArtboard(abName)) continue;

    if (proofIdx === -1 && isProofArtboard(abName)) {
      proofIdx = i;
      continue;
    }

    items.push({
      abIndex: i,
      name: abName,
      qty: parseQty(abName),
    });
  }

  if (proofIdx === -1) {
    alert('No proof artboard found. Name an artboard "PROOF" or "MOCKUP".');
    return;
  }

  if (items.length === 0) {
    alert("No print artboards found (besides PROOF/MOCKUP).");
    return;
  }

  // --- prompt details (DETAILS -> PROOF -> PRINT is irrelevant now since artboards are the source) ---
  var details = makeDetailsDialog("", "MIXER");
  if (!details) return;

  // --- temp exports ---
  var tmpFolder = Folder(Folder.temp.fsName + "/fleet_spec_sheet_ai_v1");
  if (!tmpFolder.exists) tmpFolder.create();

  var proofFile = File(tmpFolder.fsName + "/proof.png");
  exportArtboardPNG(doc, proofIdx, proofFile);

  // Export thumbnails for each print artboard
  for (var t = 0; t < items.length; t++) {
    items[t].png = File(tmpFolder.fsName + "/thumb_" + (t + 1) + ".png");
    exportArtboardPNG(doc, items[t].abIndex, items[t].png);
  }

  // --- create spec sheet artboards (pagination if needed) ---
  // Determine how many rows fit on first page (includes proof + signoff)
  function layoutCapacity(includeProof, includeSign) {
    var y = TOP;

    // header
    y -= HEADER_H;
    y -= 0.12 * IN;

    // details
    y -= DETAILS_H;
    y -= 0.14 * IN;

    // proof
    if (includeProof) {
      // We'll compute proof height later when placing; here reserve a conservative block.
      y -= 3.2 * IN; // reserved (actual proof height measured after placement)
      y -= 0.18 * IN;
    }

    // manifest title + header
    y -= MANIFEST_TITLE_GAP;
    y -= MANIFEST_HDR_H;

    // reserve signoff
    var reservedBottom = BOT;
    if (includeSign) reservedBottom += SIGN_TITLE_GAP + SIGN_H + 0.1 * IN;

    var available = y - reservedBottom;
    var rows = Math.floor(available / ROW_H);
    return rows < 0 ? 0 : rows;
  }

  // We'll do a practical approach:
  // - Page 1: header + details + proof + manifest header + rows + (signoff only if last page)
  // - Continuation pages: header + details + manifest header + rows (no proof)
  // - Last page: includes signoff under the manifest
  // Determine rows per page by iterative pagination:
  var pages = [];
  var idx = 0;

  // rough split: assume proof uses ~3" vertical; compute page 1 capacity conservative
  var firstCap = layoutCapacity(true, true);
  if (firstCap < 1) firstCap = 1;

  // We'll distribute; last page will have signoff; if more than one page, put signoff only on final.
  // Compute continuation capacity (no proof, signoff only on last)
  var contCap = layoutCapacity(false, true);
  if (contCap < 1) contCap = 1;

  // Split items
  while (idx < items.length) {
    var remaining = items.length - idx;
    var cap = pages.length === 0 ? firstCap : contCap;
    var take = Math.min(cap, remaining);
    pages.push({ start: idx, count: take });
    idx += take;
  }

  // Helper: add artboard to the right of existing doc content
  function newArtboardRectToRight() {
    // place new artboard to right of max existing bounds
    var maxRight = -1e9;
    var maxTop = 1e9;
    for (var i = 0; i < doc.artboards.length; i++) {
      var r = doc.artboards[i].artboardRect; // [left, top, right, bottom]
      if (r[2] > maxRight) maxRight = r[2];
      if (r[1] < maxTop) maxTop = r[1];
    }
    var left = maxRight + 0.5 * IN;
    var top = doc.artboards[0].artboardRect[1]; // align top to first artboard top
    return [left, top, left + PAGE_W, top - PAGE_H];
  }

  function drawSpecPage(layer, pageNum, isFirst, isLast, chunk) {
    var y = TOP;
    var x = X0;

    // Header bar
    var headerRect = addRect(layer, x, y, CONTENT_W, HEADER_H, true);
    var headerTF = addText(
      layer,
      x + 0.18 * IN,
      y - 0.12 * IN,
      CONTENT_W - 0.36 * IN,
      HEADER_H,
      HEADER_TEXT,
      14,
      [255, 255, 255],
      HUDSON
    );

    y -= HEADER_H + 0.12 * IN;

    // Details table (2 rows, fixed width = 7.5")
    var tableW = CONTENT_W;
    var rowH = DETAILS_H / 2;

    // outer box
    addRect(layer, x, y, tableW, DETAILS_H, false);

    // horizontal divider
    addLine(layer, x, y - rowH, x + tableW, y - rowH);

    // column plan: label + big value + label + unit label + unit blank (spanned like we set earlier)
    // We'll mimic the latest 2-row approach:
    // Row1: CUSTOMER: | customer spans across
    // Row2: VEHICLE:  | vehicle spans | VEHICLE UNIT #: | blank
    var c1 = 1.2 * IN,
      c2 = 4.4 * IN,
      c5 = 1.3 * IN,
      c6 = 0.6 * IN;
    // vertical lines
    addLine(layer, x + c1, y, x + c1, y - DETAILS_H); // after first label col
    addLine(layer, x + c1 + c2, y - rowH, x + c1 + c2, y - DETAILS_H); // only row2 split point (vehicle span ends early)
    addLine(
      layer,
      x + tableW - (c5 + c6),
      y - rowH,
      x + tableW - (c5 + c6),
      y - DETAILS_H
    );
    addLine(layer, x + tableW - c6, y - rowH, x + tableW - c6, y - DETAILS_H);

    // Row 1 texts
    addText(
      layer,
      x + 0.08 * IN,
      y - 0.14 * IN,
      c1 - 0.16 * IN,
      rowH,
      "CUSTOMER:",
      9,
      [0, 0, 0],
      HUDSON
    );
    addText(
      layer,
      x + c1 + 0.08 * IN,
      y - 0.14 * IN,
      tableW - c1 - 0.16 * IN,
      rowH,
      details.customer,
      9,
      [0, 0, 0],
      HUDSON
    );

    // Row 2 texts
    var y2 = y - rowH;
    addText(
      layer,
      x + 0.08 * IN,
      y2 - 0.14 * IN,
      c1 - 0.16 * IN,
      rowH,
      "VEHICLE:",
      9,
      [0, 0, 0],
      HUDSON
    );
    addText(
      layer,
      x + c1 + 0.08 * IN,
      y2 - 0.14 * IN,
      c2 - 0.16 * IN,
      rowH,
      details.vehicle,
      9,
      [0, 0, 0],
      HUDSON
    );
    addText(
      layer,
      x + tableW - (c5 + c6) + 0.08 * IN,
      y2 - 0.14 * IN,
      c5 - 0.16 * IN,
      rowH,
      "VEHICLE UNIT #:",
      9,
      [0, 0, 0],
      HUDSON
    );
    // blank box is last col area; leave empty on purpose

    y -= DETAILS_H + 0.18 * IN;

    // Proof (only on first page)
    if (isFirst) {
      // Outer proof border box; place image centered inside with inner width 7.2"
      addRect(layer, x, y, PROOF_OUTER_W, 0.01 * IN, false); // temp, will resize after placing

      var placedProof = layer.placedItems.add();
      placedProof.file = proofFile;

      // Fit to inner width (7.2") and keep aspect; then build box around it with width 7.5"
      // Determine current size
      fitPlacedToBox(placedProof, PROOF_INNER_W, 1000 * IN);

      // Center within outer width (7.5") with small side padding
      var vbP = placedProof.visibleBounds;
      var pW = vbP[2] - vbP[0];
      var pH = vbP[1] - vbP[3];

      var proofTop = y;
      placedProof.left = toDocX(x + (PROOF_OUTER_W - pW) / 2);
      placedProof.top = toDocY(proofTop);

      // Make actual border box height = image height
      var proofBoxH = pH;
      // remove temp box and replace with correct
      // (Illustrator doesn't let us resize a pathItem.rectangle easily by setting height? We'll recreate.)
      // Border box sized to actual placed proof height
      var proofBorder = addRect(
        layer,
        x,
        proofTop,
        PROOF_OUTER_W,
        proofBoxH,
        false
      );

      y = proofTop - proofBoxH - 0.18 * IN;
    }

    // PRINT MANIFEST title (underlined by drawing a line)
    var titleTF = addText(
      layer,
      x,
      y,
      CONTENT_W,
      0.25 * IN,
      "PRINT MANIFEST",
      11,
      [0, 0, 0],
      HUDSON
    );
    // underline
    addLine(layer, x, y - 0.2 * IN, x + 1.7 * IN, y - 0.2 * IN);

    y -= MANIFEST_TITLE_GAP;

    // Manifest table header row
    var tableX = x;
    var tableY = y;

    var tableW2 = CONTENT_W;
    var hdrH = MANIFEST_HDR_H;

    // header row box
    addRect(layer, tableX, tableY, tableW2, hdrH, false);

    // vertical lines for columns
    var cx = tableX;
    var cols = [COL_W.thumb, COL_W.name, COL_W.file, COL_W.qty, COL_W.done];
    for (var c = 0; c < cols.length - 1; c++) {
      cx += cols[c];
      addLine(layer, cx, tableY, cx, tableY - hdrH);
    }

    // header texts
    var hx = tableX;
    addText(
      layer,
      hx + 0.08 * IN,
      tableY - 0.12 * IN,
      cols[0] - 0.16 * IN,
      hdrH,
      "THUMBNAIL",
      9,
      [0, 0, 0],
      HUDSON
    );
    hx += cols[0];
    addText(
      layer,
      hx + 0.08 * IN,
      tableY - 0.12 * IN,
      cols[1] - 0.16 * IN,
      hdrH,
      "NAME",
      9,
      [0, 0, 0],
      HUDSON
    );
    hx += cols[1];
    addText(
      layer,
      hx + 0.08 * IN,
      tableY - 0.12 * IN,
      cols[2] - 0.16 * IN,
      hdrH,
      "FILE NAME",
      9,
      [0, 0, 0],
      HUDSON
    );
    hx += cols[2];
    addText(
      layer,
      hx + 0.1 * IN,
      tableY - 0.12 * IN,
      cols[3] - 0.2 * IN,
      hdrH,
      "QTY",
      9,
      [0, 0, 0],
      HUDSON
    );
    hx += cols[3];
    addText(
      layer,
      hx + 0.08 * IN,
      tableY - 0.12 * IN,
      cols[4] - 0.16 * IN,
      hdrH,
      "DONE",
      9,
      [0, 0, 0],
      HUDSON
    );

    y = tableY - hdrH;

    // body rows
    for (var r = 0; r < chunk.count; r++) {
      var item = items[chunk.start + r];
      var rowTop = y;

      // row outer rect
      addRect(layer, tableX, rowTop, tableW2, ROW_H, false);

      // vertical lines
      cx = tableX;
      for (c = 0; c < cols.length - 1; c++) {
        cx += cols[c];
        addLine(layer, cx, rowTop, cx, rowTop - ROW_H);
      }

      // thumbnail
      var thumbPlaced = layer.placedItems.add();
      thumbPlaced.file = item.png;
      fitPlacedToBox(thumbPlaced, THUMB_BOX_W, THUMB_BOX_H);
      centerPlacedInCell(thumbPlaced, tableX, rowTop, cols[0], ROW_H);

      // NAME + FILE + QTY text (top aligned feel: set top slightly below rowTop)
      var textTop = rowTop - 0.14 * IN;
      var nameText = (item.name || "").toUpperCase();
      var fileText = (item.name || "").toUpperCase(); // no folder paths in AI mode
      var qtyText = String(item.qty);

      var tx = tableX + cols[0];
      addText(
        layer,
        tx + 0.08 * IN,
        textTop,
        cols[1] - 0.16 * IN,
        ROW_H,
        nameText,
        9,
        [0, 0, 0],
        HUDSON
      );
      tx += cols[1];
      addText(
        layer,
        tx + 0.08 * IN,
        textTop,
        cols[2] - 0.16 * IN,
        ROW_H,
        fileText,
        9,
        [0, 0, 0],
        HUDSON
      );
      tx += cols[2];
      addText(
        layer,
        tx + 0.1 * IN,
        textTop,
        cols[3] - 0.2 * IN,
        ROW_H,
        qtyText,
        9,
        [0, 0, 0],
        HUDSON
      );

      y -= ROW_H;
    }

    // Sign-off only on last page
    if (isLast) {
      y -= SIGN_TITLE_GAP;

      var signTitle = addText(
        layer,
        x,
        y,
        CONTENT_W,
        0.25 * IN,
        "INSTALL SIGN-OFF",
        11,
        [0, 0, 0],
        HUDSON
      );
      addLine(layer, x, y - 0.2 * IN, x + 1.75 * IN, y - 0.2 * IN);

      y -= 0.28 * IN;

      // sign table 2 rows
      var signW = CONTENT_W;
      var signX = x;
      var signTop = y;

      var widths = [1.3 * IN, 2.2 * IN, 1.1 * IN, 0.9 * IN, 0.7 * IN, 1.3 * IN];
      var signRow1 = 0.35 * IN;
      var signRow2 = 1.2 * IN;
      var signH = signRow1 + signRow2;

      addRect(layer, signX, signTop, signW, signH, false);
      addLine(
        layer,
        signX,
        signTop - signRow1,
        signX + signW,
        signTop - signRow1
      );

      // verticals row1
      cx = signX;
      for (c = 0; c < widths.length - 1; c++) {
        cx += widths[c];
        addLine(layer, cx, signTop, cx, signTop - signRow1);
      }
      // verticals row2: only first column; rest is spanned
      addLine(
        layer,
        signX + widths[0],
        signTop - signRow1,
        signX + widths[0],
        signTop - signH
      );

      // row1 text
      var sx = signX;
      addText(
        layer,
        sx + 0.08 * IN,
        signTop - 0.11 * IN,
        widths[0] - 0.16 * IN,
        signRow1,
        "INSTALLED BY:",
        9,
        [0, 0, 0],
        HUDSON
      );
      sx += widths[0];
      // blank cell (no underscores)
      sx += widths[1];
      addText(
        layer,
        sx + 0.08 * IN,
        signTop - 0.11 * IN,
        widths[2] - 0.16 * IN,
        signRow1,
        "PHOTOS TAKEN:",
        9,
        [0, 0, 0],
        HUDSON
      );
      sx += widths[2];
      addText(
        layer,
        sx + 0.1 * IN,
        signTop - 0.11 * IN,
        widths[3] - 0.2 * IN,
        signRow1,
        "YES   NO",
        9,
        [0, 0, 0],
        HUDSON
      );
      sx += widths[3];
      addText(
        layer,
        sx + 0.08 * IN,
        signTop - 0.11 * IN,
        widths[4] - 0.16 * IN,
        signRow1,
        "DATE:",
        9,
        [0, 0, 0],
        HUDSON
      );
      sx += widths[4];
      // blank cell (no underscores)

      // row2 label
      addText(
        layer,
        signX + 0.08 * IN,
        signTop - signRow1 - 0.14 * IN,
        widths[0] - 0.16 * IN,
        signRow2,
        "ISSUES / DAMAGE NOTES:",
        9,
        [0, 0, 0],
        HUDSON
      );
      // rest is blank by design
    }
  }

  // Create pages/artboards + layers
  var createdArtboards = [];
  var firstNewIdx = doc.artboards.length;

  for (var p = 0; p < pages.length; p++) {
    var rect = newArtboardRectToRight();
    doc.artboards.add(rect);
    var abIdx = doc.artboards.length - 1;
    doc.artboards[abIdx].name =
      p === 0 ? "SPEC SHEET (V1)" : "SPEC SHEET (V1) — Pg " + (p + 1);
    createdArtboards.push(abIdx);

    // Make a new layer for each page for cleanliness
    var layer = doc.layers.add();
    layer.name = doc.artboards[abIdx].name;

    // Activate artboard so coordinates are aligned to it
    doc.artboards.setActiveArtboardIndex(abIdx);

    // translate everything to this artboard's coordinate space
    // We'll work in artboard-local by offsetting with artboardRect left/top.
    var abRect = doc.artboards[abIdx].artboardRect; // [L, T, R, B]
    __AB_LEFT = abRect[0];
    __AB_TOP = abRect[1];
    // Create a group and shift by artboard origin
    var g = layer.groupItems.add();
    g.left = abRect[0];
    g.top = abRect[1];

    // Draw into group as "layer" stand-in by temporarily creating items on layer then moving
    // Simpler: draw directly on layer but offset x,y by abRect[0],abRect[1]
    // We'll just add offsets:
    var oldAddRect = addRect;
    var oldAddLine = addLine;
    var oldAddText = addText;

    function addRectO(layerRef, x, yTop, w, h, fill) {
      return oldAddRect(
        layerRef,
        abRect[0] + x,
        abRect[1] - (PAGE_H - yTop),
        w,
        h,
        fill
      );
    }
    function addLineO(layerRef, x1, y1, x2, y2) {
      return oldAddLine(
        layerRef,
        abRect[0] + x1,
        abRect[1] - (PAGE_H - y1),
        abRect[0] + x2,
        abRect[1] - (PAGE_H - y2)
      );
    }
    function addTextO(layerRef, x, yTop, w, h, s, size, rgb, fontObj, align) {
      return oldAddText(
        layerRef,
        abRect[0] + x,
        abRect[1] - (PAGE_H - yTop),
        w,
        h,
        s,
        size,
        rgb,
        fontObj,
        align
      );
    }
    // override globals used by drawSpecPage
    addRect = addRectO;
    addLine = addLineO;
    addText = addTextO;

    drawSpecPage(layer, p + 1, p === 0, p === pages.length - 1, pages[p]);

    // restore
    addRect = oldAddRect;
    addLine = oldAddLine;
    addText = oldAddText;
  }

  alert("Created " + pages.length + " SPEC SHEET (V1) artboard(s).");
})();
