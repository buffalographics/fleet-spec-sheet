/* 
Fleet Spec Sheet Generator — Illustrator (V1)

V1 improvements:
- No temp export PNGs / no tmp layers.
- Build thumbnails + proof by duplicating artwork from source artboards directly.
- Rasterize any linked/placed artwork used in the spec sheet.
- Update in-place: re-run updates the existing SPEC SHEET (V1) artboard + layer.
- Pagination dropped: always expand spec-sheet artboard height so nothing hangs off.

Artboard rules:
- Proof artboard name contains: PROOF or MOCKUP (case-insensitive)
- Ignored artboards:
  - name contains 'unit' anywhere (case-insensitive)
  - OR name is pure digits (e.g. '4519')
- Qty parsing: QTY2, QTY_2, QTY-2, QTY 2, x2, 2x
*/

(function () {
  if (app.documents.length === 0) return;
  var doc = app.activeDocument;

  var IN = 72;
  var PAGE_W = 8.5 * IN;

  var MARGIN = 0.25 * IN;
  var X0 = MARGIN;
  var CONTENT_W = PAGE_W - 2 * MARGIN;

  var HEADER_H = 0.45 * IN;
  var HEADER_TEXT = "BUFFALO GRAPHICS COMPANY";

  var DETAILS_H = 0.6 * IN;

  var PROOF_OUTER_W = CONTENT_W;
  var PROOF_INNER_W = CONTENT_W - 0.3 * IN;

  var MANIFEST_TITLE_H = 0.25 * IN;
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

  var SIGN_TITLE_GAP = 0.2 * IN;
  var SIGN_TITLE_H = 0.25 * IN;
  var SIGN_TABLE_GAP = 0.28 * IN;
  var SIGN_H = 1.55 * IN;

  function trim(s) {
    return String(s || "").replace(/^\s+|\s+$/g, "");
  }

  function tryFont(name) {
    try {
      return app.textFonts.getByName(name);
    } catch (e) {
      return null;
    }
  }
  function findFontContains(substr) {
    try {
      var needle = String(substr || "").toLowerCase();
      for (var i = 0; i < app.textFonts.length; i++) {
        var f = app.textFonts[i];
        var n = f && f.name ? String(f.name).toLowerCase() : "";
        var fam = f && f.family ? String(f.family).toLowerCase() : "";
        if (n.indexOf(needle) !== -1 || fam.indexOf(needle) !== -1) return f;
      }
    } catch (e) {}
    return null;
  }
  var HUDSON =
    tryFont("HudsonNY") ||
    tryFont("Hudson NY") ||
    tryFont("HudsonNY-Regular") ||
    findFontContains("hudson");

  function setText(tf, size, rgb, fontObj) {
    var tr = tf.textRange;
    tr.characterAttributes.size = size;
    if (fontObj) tr.characterAttributes.textFont = fontObj;
    else if (HUDSON) tr.characterAttributes.textFont = HUDSON;
    if (rgb) {
      var c = new RGBColor();
      c.red = rgb[0];
      c.green = rgb[1];
      c.blue = rgb[2];
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

  function addRect(container, left, top, w, h, fill) {
    var r = container.pathItems.rectangle(top, left, w, h);
    if (fill) fillBlack(r);
    else strokeBlack(r, 1);
    return r;
  }
  function addLine(container, x1, y1, x2, y2) {
    var ln = container.pathItems.add();
    ln.setEntirePath([
      [x1, y1],
      [x2, y2],
    ]);
    strokeBlack(ln, 1);
    return ln;
  }
  function addAreaText(
    container,
    left,
    top,
    w,
    h,
    s,
    size,
    rgb,
    fontObj,
    justify,
  ) {
    var box = container.pathItems.rectangle(top, left, w, h);
    box.stroked = false;
    box.filled = false;
    var tf = container.textFrames.areaText(box);
    tf.contents = s;
    setText(tf, size, rgb, fontObj);
    if (justify) {
      try {
        tf.textRange.paragraphAttributes.justification = justify;
      } catch (e) {}
    }
    return tf;
  }

  function parseQty(name) {
    var n = String(name || "");
    var m;
    m = n.match(/(?:^|[_\-\s])qty(?:[_\-\s]*)(\d+)(?:[_\-\s\.]|$)/i);
    if (m) return parseInt(m[1], 10);
    m = n.match(/(?:^|[_\-\s])x(\d+)(?:[_\-\s\.]|$)/i);
    if (m) return parseInt(m[1], 10);
    m = n.match(/(?:^|[_\-\s])(\d+)x(?:[_\-\s\.]|$)/i);
    if (m) return parseInt(m[1], 10);
    return 1;
  }

  function isProofArtboard(name) {
    var n = String(name || "");
    return /proof/i.test(n) || /mockup/i.test(n);
  }
  function isIgnoredArtboard(name) {
    if (!name) return false;
    var n = trim(String(name));
    n = n.replace(/^\s*\d+\s*([\-\._]|\u2013|\u2014)\s*/g, "");
    n = n.replace(/^\s*\d+\s+/, "");
    if (/unit/i.test(n)) return true;
    if (/^\d+$/.test(n)) return true;
    return false;
  }

  function getClientsSegmentsFromDocPath(docRef) {
    try {
      var f = null;
      try {
        f = docRef.fullName;
      } catch (e0) {
        f = null;
      }
      if (!f) return null;
      var norm = String(f.fsName || f.fullName || "").replace(/\\/g, "/");
      var parts = norm.split("/");
      var idx = -1;
      for (var i = 0; i < parts.length; i++) {
        if (String(parts[i]).toLowerCase() === "clients") {
          idx = i;
          break;
        }
      }
      if (idx === -1) return null;
      return {
        customer: trim(parts[idx + 1] || ""),
        job: trim(parts[idx + 2] || ""),
      };
    } catch (e) {
      return null;
    }
  }
  function inferCustomerFromDocPath(docRef) {
    var seg = getClientsSegmentsFromDocPath(docRef);
    return seg && seg.customer ? seg.customer : "";
  }
  function inferVehicleFromDocPath(docRef) {
    var seg = getClientsSegmentsFromDocPath(docRef);
    if (seg && seg.job) return seg.job;
    try {
      var f = null;
      try {
        f = docRef.fullName;
      } catch (e0) {
        f = null;
      }
      if (!f) return "";
      var norm = String(f.fsName || f.fullName || "")
        .replace(/\\/g, "/")
        .toLowerCase();
      if (/\bmixer\b|\bmixers\b|ready\s*mix|\brm\b/.test(norm)) return "MIXER";
      if (/\bpump\b|\bboom\b/.test(norm)) return "PUMP";
      if (/\btrailer\b/.test(norm)) return "TRAILER";
      if (/\bdump\b/.test(norm)) return "DUMP";
      return "";
    } catch (e) {
      return "";
    }
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
    btns.add("button", undefined, "OK", { name: "ok" });

    customer.active = true;
    if (w.show() !== 1) return null;

    return {
      customer: trim(customer.text),
      vehicle: trim(vehicle.text) || "MIXER",
    };
  }

  function collectPageItemsInArtboard(abIndex, excludeLayerName) {
    var ab = doc.artboards[abIndex];
    var r = ab.artboardRect; // [L, T, R, B]
    var L = r[0],
      T = r[1],
      R = r[2],
      B = r[3];

    var res = [];
    for (var i = 0; i < doc.pageItems.length; i++) {
      var it = doc.pageItems[i];
      try {
        if (!it || it.locked || it.hidden) continue;
        if (excludeLayerName && it.layer && it.layer.name === excludeLayerName)
          continue;

        var b = it.visibleBounds; // [l,t,r,b]
        if (!b || b.length !== 4) continue;

        var ol = Math.max(L, b[0]);
        var ot = Math.min(T, b[1]);
        var orr = Math.min(R, b[2]);
        var ob = Math.max(B, b[3]);

        if (orr > ol && ot > ob) res.push(it);
      } catch (e) {}
    }
    return res;
  }

  function duplicateArtboardArtworkToGroup(
    abIndex,
    targetGroup,
    excludeLayerName,
  ) {
    var arr = collectPageItemsInArtboard(abIndex, excludeLayerName);
    for (var i = 0; i < arr.length; i++) {
      try {
        arr[i].duplicate(targetGroup, ElementPlacement.PLACEATEND);
      } catch (e) {}
    }
  }

  function fitGroupToBox(group, boxW, boxH) {
    var vb = group.visibleBounds; // [L,T,R,B]
    var w = vb[2] - vb[0];
    var h = vb[1] - vb[3];
    if (w <= 0 || h <= 0) return;

    var s = Math.min(boxW / w, boxH / h) * 100;
    group.resize(s, s, true, true, true, true, s, Transformation.CENTER);
  }

  function centerGroupInBox(group, boxLeft, boxTop, boxW, boxH) {
    var vb = group.visibleBounds;
    var w = vb[2] - vb[0];
    var h = vb[1] - vb[3];

    var cx = boxLeft + boxW / 2;
    var cy = boxTop - boxH / 2;

    var curL = vb[0];
    var curT = vb[1];

    var targetL = cx - w / 2;
    var targetT = cy + h / 2;

    group.translate(targetL - curL, targetT - curT);
  }

  function rasterizeItem(item) {
    try {
      if (!item) return null;
      var vb = item.visibleBounds;
      if (!vb || vb.length !== 4) return null;

      var ro = new RasterizeOptions();
      ro.resolution = 300;
      ro.antiAliasingMethod = AntiAliasingMethod.ARTOPTIMIZED;
      ro.transparency = true;
      ro.backgroundBlack = false;
      try {
        ro.convertSpotColors = false;
      } catch (e0) {}

      var r = doc.rasterize(item, vb, ro);
      try {
        item.remove();
      } catch (e1) {}
      return r;
    } catch (e) {
      return null;
    }
  }

  function getOrCreateSpecArtboard() {
    for (var i = 0; i < doc.artboards.length; i++) {
      if (String(doc.artboards[i].name || "").indexOf("SPEC SHEET (V1)") === 0)
        return i;
    }

    var maxRight = -1e12,
      topAlign = doc.artboards[0].artboardRect[1];
    for (var j = 0; j < doc.artboards.length; j++) {
      var rr = doc.artboards[j].artboardRect;
      if (rr[2] > maxRight) maxRight = rr[2];
    }
    var gap = 0.5 * IN;
    var left = maxRight + gap;
    var top = topAlign;
    var rect = [left, top, left + PAGE_W, top - 11 * IN];
    var idx = doc.artboards.length;
    try {
      doc.artboards.add(rect);
      idx = doc.artboards.length - 1;
    } catch (e1) {
      var minBottom = 1e12,
        leftAlign = doc.artboards[0].artboardRect[0];
      for (var k = 0; k < doc.artboards.length; k++) {
        var rr2 = doc.artboards[k].artboardRect;
        if (rr2[3] < minBottom) minBottom = rr2[3];
      }
      var top2 = minBottom - gap;
      var rect2 = [leftAlign, top2, leftAlign + PAGE_W, top2 - 11 * IN];
      doc.artboards.add(rect2);
      idx = doc.artboards.length - 1;
    }
    try {
      doc.artboards[idx].name = "SPEC SHEET (V1)";
    } catch (eName) {}
    return idx;
  }

  function getOrCreateSpecLayer() {
    for (var i = 0; i < doc.layers.length; i++) {
      if (doc.layers[i].name === "SPEC SHEET (V1)") return doc.layers[i];
    }
    var lyr = doc.layers.add();
    lyr.name = "SPEC SHEET (V1)";
    return lyr;
  }

  function clearLayer(layer) {
    try {
      for (var i = layer.pageItems.length - 1; i >= 0; i--) {
        try {
          layer.pageItems[i].remove();
        } catch (e1) {}
      }
      for (var j = layer.layers.length - 1; j >= 0; j--) {
        try {
          layer.layers[j].remove();
        } catch (e2) {}
      }
    } catch (e) {}
  }

  function expandArtboardToFitContent(abIdx, contentGroup) {
    try {
      var ab = doc.artboards[abIdx];
      if (!ab) return;
      var r = ab.artboardRect;
      var b = contentGroup.visibleBounds;
      if (!b || b.length !== 4) return;

      var pad = 0.25 * IN;
      var newL = r[0];
      var newR = r[0] + PAGE_W;
      var newT = r[1];
      var newB = Math.min(r[3], b[3] - pad);

      ab.artboardRect = [newL, newT, newR, newB];
    } catch (e) {}
  }

  // Gather artboards
  var proofIdx = -1;
  var items = [];
  for (var i = 0; i < doc.artboards.length; i++) {
    var abName = doc.artboards[i].name || "Artboard " + (i + 1);
    if (isIgnoredArtboard(abName)) continue;

    if (proofIdx === -1 && isProofArtboard(abName)) {
      proofIdx = i;
      continue;
    }

    items.push({ abIndex: i, name: abName, qty: parseQty(abName) });
  }
  if (proofIdx === -1 || items.length === 0) return;

  // Details prompt
  var inferredCustomer = inferCustomerFromDocPath(doc);
  var inferredVehicle = inferVehicleFromDocPath(doc) || "MIXER";
  var details = makeDetailsDialog(inferredCustomer, inferredVehicle);
  if (!details) return;

  // Build/update
  var specAbIdx = getOrCreateSpecArtboard();
  var specLayer = getOrCreateSpecLayer();
  clearLayer(specLayer);

  doc.artboards.setActiveArtboardIndex(specAbIdx);
  var abRect = doc.artboards[specAbIdx].artboardRect;
  var AB_L = abRect[0];
  var AB_T = abRect[1];

  var root = specLayer.groupItems.add();
  root.name = "__spec_sheet_root__";

  var y = AB_T;
  var headerTop = AB_T;
  var pageLeft = AB_L;

  // Header full width, flush top
  addRect(root, pageLeft, headerTop, PAGE_W, HEADER_H, true);
  addAreaText(
    root,
    pageLeft + 0.18 * IN,
    headerTop - 0.12 * IN,
    PAGE_W - 0.36 * IN,
    HEADER_H,
    HEADER_TEXT,
    14,
    [255, 255, 255],
    HUDSON,
  );

  y = headerTop - HEADER_H - 0.12 * IN;

  // Details
  var tableLeft = pageLeft + X0;
  var tableTop = y;
  var tableW = CONTENT_W;
  var rowH = DETAILS_H / 2;

  addRect(root, tableLeft, tableTop, tableW, DETAILS_H, false);
  addLine(
    root,
    tableLeft,
    tableTop - rowH,
    tableLeft + tableW,
    tableTop - rowH,
  );

  var c1 = 1.2 * IN,
    c2 = 4.4 * IN,
    c5 = 1.3 * IN,
    c6 = 0.6 * IN;
  addLine(root, tableLeft + c1, tableTop, tableLeft + c1, tableTop - DETAILS_H);
  addLine(
    root,
    tableLeft + c1 + c2,
    tableTop - rowH,
    tableLeft + c1 + c2,
    tableTop - DETAILS_H,
  );
  addLine(
    root,
    tableLeft + tableW - (c5 + c6),
    tableTop - rowH,
    tableLeft + tableW - (c5 + c6),
    tableTop - DETAILS_H,
  );
  addLine(
    root,
    tableLeft + tableW - c6,
    tableTop - rowH,
    tableLeft + tableW - c6,
    tableTop - DETAILS_H,
  );

  addAreaText(
    root,
    tableLeft + 0.08 * IN,
    tableTop - 0.14 * IN,
    c1 - 0.16 * IN,
    rowH,
    "CUSTOMER:",
    9,
    [0, 0, 0],
    HUDSON,
  );
  addAreaText(
    root,
    tableLeft + c1 + 0.08 * IN,
    tableTop - 0.14 * IN,
    tableW - c1 - 0.16 * IN,
    rowH,
    details.customer,
    9,
    [0, 0, 0],
    HUDSON,
  );

  var row2Top = tableTop - rowH;
  addAreaText(
    root,
    tableLeft + 0.08 * IN,
    row2Top - 0.14 * IN,
    c1 - 0.16 * IN,
    rowH,
    "VEHICLE:",
    9,
    [0, 0, 0],
    HUDSON,
  );
  addAreaText(
    root,
    tableLeft + c1 + 0.08 * IN,
    row2Top - 0.14 * IN,
    c2 - 0.16 * IN,
    rowH,
    details.vehicle,
    9,
    [0, 0, 0],
    HUDSON,
  );
  addAreaText(
    root,
    tableLeft + tableW - (c5 + c6) + 0.08 * IN,
    row2Top - 0.14 * IN,
    c5 - 0.16 * IN,
    rowH,
    "VEHICLE UNIT #:",
    9,
    [0, 0, 0],
    HUDSON,
  );

  y = tableTop - DETAILS_H - 0.18 * IN;

  // Proof
  var proofTop = y;
  var proofGroup = root.groupItems.add();
  proofGroup.name = "__proof_group__";
  duplicateArtboardArtworkToGroup(proofIdx, proofGroup, "SPEC SHEET (V1)");
  fitGroupToBox(proofGroup, PROOF_INNER_W, 1000 * IN);

  var pvb = proofGroup.visibleBounds;
  var pW = pvb[2] - pvb[0];
  var pH = pvb[1] - pvb[3];

  var proofBoxLeft = pageLeft + X0;
  var proofInnerLeft = proofBoxLeft + (PROOF_OUTER_W - pW) / 2;
  proofGroup.translate(proofInnerLeft - pvb[0], proofTop - pvb[1]);
  rasterizeItem(proofGroup);
  addRect(root, proofBoxLeft, proofTop, PROOF_OUTER_W, pH, false);

  y = proofTop - pH - 0.18 * IN;

  // Manifest title
  addAreaText(
    root,
    pageLeft + X0,
    y,
    CONTENT_W,
    MANIFEST_TITLE_H,
    "PRINT MANIFEST",
    11,
    [0, 0, 0],
    HUDSON,
  );
  addLine(
    root,
    pageLeft + X0,
    y - 0.2 * IN,
    pageLeft + X0 + 1.7 * IN,
    y - 0.2 * IN,
  );
  y -= MANIFEST_TITLE_H;
  y -= MANIFEST_TITLE_GAP;

  // Manifest header row
  var tX = pageLeft + X0;
  var tY = y;
  var tW = CONTENT_W;

  addRect(root, tX, tY, tW, MANIFEST_HDR_H, false);

  var cols = [COL_W.thumb, COL_W.name, COL_W.file, COL_W.qty, COL_W.done];
  var cx = tX;
  for (var c = 0; c < cols.length - 1; c++) {
    cx += cols[c];
    addLine(root, cx, tY, cx, tY - MANIFEST_HDR_H);
  }

  var hx = tX;
  addAreaText(
    root,
    hx + 0.08 * IN,
    tY - 0.12 * IN,
    cols[0] - 0.16 * IN,
    MANIFEST_HDR_H,
    "THUMBNAIL",
    9,
    [0, 0, 0],
    HUDSON,
  );
  hx += cols[0];
  addAreaText(
    root,
    hx + 0.08 * IN,
    tY - 0.12 * IN,
    cols[1] - 0.16 * IN,
    MANIFEST_HDR_H,
    "NAME",
    9,
    [0, 0, 0],
    HUDSON,
  );
  hx += cols[1];
  addAreaText(
    root,
    hx + 0.08 * IN,
    tY - 0.12 * IN,
    cols[2] - 0.16 * IN,
    MANIFEST_HDR_H,
    "FILE NAME",
    9,
    [0, 0, 0],
    HUDSON,
  );
  hx += cols[2];
  addAreaText(
    root,
    hx + 0.1 * IN,
    tY - 0.12 * IN,
    cols[3] - 0.2 * IN,
    MANIFEST_HDR_H,
    "QTY",
    9,
    [0, 0, 0],
    HUDSON,
  );
  hx += cols[3];
  addAreaText(
    root,
    hx + 0.08 * IN,
    tY - 0.12 * IN,
    cols[4] - 0.16 * IN,
    MANIFEST_HDR_H,
    "DONE",
    9,
    [0, 0, 0],
    HUDSON,
  );

  y = tY - MANIFEST_HDR_H;

  // Rows
  for (var r = 0; r < items.length; r++) {
    var it = items[r];
    var rowTop = y;

    addRect(root, tX, rowTop, tW, ROW_H, false);

    cx = tX;
    for (c = 0; c < cols.length - 1; c++) {
      cx += cols[c];
      addLine(root, cx, rowTop, cx, rowTop - ROW_H);
    }

    var thumbGroup = root.groupItems.add();
    thumbGroup.name = "__thumb_" + (r + 1) + "__";
    duplicateArtboardArtworkToGroup(it.abIndex, thumbGroup, "SPEC SHEET (V1)");
    fitGroupToBox(thumbGroup, THUMB_BOX_W, THUMB_BOX_H);
    centerGroupInBox(thumbGroup, tX, rowTop, cols[0], ROW_H);
    rasterizeItem(thumbGroup);

    var textTop = rowTop - 0.14 * IN;
    var nameText = String(it.name || "").toUpperCase();
    var fileText = String(it.name || "").toUpperCase();
    var qtyText = String(it.qty);

    var tx = tX + cols[0];
    addAreaText(
      root,
      tx + 0.08 * IN,
      textTop,
      cols[1] - 0.16 * IN,
      ROW_H,
      nameText,
      9,
      [0, 0, 0],
      HUDSON,
    );
    tx += cols[1];
    addAreaText(
      root,
      tx + 0.08 * IN,
      textTop,
      cols[2] - 0.16 * IN,
      ROW_H,
      fileText,
      9,
      [0, 0, 0],
      HUDSON,
    );
    tx += cols[2];
    addAreaText(
      root,
      tx + 0.1 * IN,
      textTop,
      cols[3] - 0.2 * IN,
      ROW_H,
      qtyText,
      9,
      [0, 0, 0],
      HUDSON,
    );

    y -= ROW_H;
  }

  // Sign-off
  y -= SIGN_TITLE_GAP;
  addAreaText(
    root,
    pageLeft + X0,
    y,
    CONTENT_W,
    SIGN_TITLE_H,
    "INSTALL SIGN-OFF",
    11,
    [0, 0, 0],
    HUDSON,
  );
  addLine(
    root,
    pageLeft + X0,
    y - 0.2 * IN,
    pageLeft + X0 + 1.75 * IN,
    y - 0.2 * IN,
  );
  y -= SIGN_TITLE_H;
  y -= SIGN_TABLE_GAP;

  var signX = pageLeft + X0;
  var signTop = y;
  var signW = CONTENT_W;

  var widths = [1.3 * IN, 2.2 * IN, 1.1 * IN, 0.9 * IN, 0.7 * IN, 1.3 * IN];
  var signRow1 = 0.35 * IN;
  var signRow2 = 1.2 * IN;
  var signH = signRow1 + signRow2;

  addRect(root, signX, signTop, signW, signH, false);
  addLine(root, signX, signTop - signRow1, signX + signW, signTop - signRow1);

  cx = signX;
  for (c = 0; c < widths.length - 1; c++) {
    cx += widths[c];
    addLine(root, cx, signTop, cx, signTop - signRow1);
  }
  addLine(
    root,
    signX + widths[0],
    signTop - signRow1,
    signX + widths[0],
    signTop - signH,
  );

  var sx = signX;
  addAreaText(
    root,
    sx + 0.08 * IN,
    signTop - 0.11 * IN,
    widths[0] - 0.16 * IN,
    signRow1,
    "INSTALLED BY:",
    9,
    [0, 0, 0],
    HUDSON,
  );
  sx += widths[0] + widths[1];
  addAreaText(
    root,
    sx + 0.08 * IN,
    signTop - 0.11 * IN,
    widths[2] - 0.16 * IN,
    signRow1,
    "PHOTOS TAKEN:",
    9,
    [0, 0, 0],
    HUDSON,
  );
  sx += widths[2];
  addAreaText(
    root,
    sx + 0.1 * IN,
    signTop - 0.11 * IN,
    widths[3] - 0.2 * IN,
    signRow1,
    "YES   NO",
    9,
    [0, 0, 0],
    HUDSON,
  );
  sx += widths[3];
  addAreaText(
    root,
    sx + 0.08 * IN,
    signTop - 0.11 * IN,
    widths[4] - 0.16 * IN,
    signRow1,
    "DATE:",
    9,
    [0, 0, 0],
    HUDSON,
  );

  addAreaText(
    root,
    signX + 0.08 * IN,
    signTop - signRow1 - 0.14 * IN,
    widths[0] - 0.16 * IN,
    signRow2,
    "ISSUES / DAMAGE NOTES:",
    9,
    [0, 0, 0],
    HUDSON,
  );

  // Critical: expand the artboard so nothing hangs off the bottom
  expandArtboardToFitContent(specAbIdx, root);

  try {
    specLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
  } catch (eZ) {}
})();
