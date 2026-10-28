// Fills Ayvens' own "Skademelding-motorvogn" PDF (skademelding-ayvens.pdf) with pdf-lib.
// data.text:   { "Text20": "AB12345", … }      – text fields by PDF field name
// data.checks: ["Check Box78", …]              – check boxes to tick
// data.choices:{ "NEIJA1": 0 | 1 }              – NEI/JA (or A/B, Kvinne/Mann) pairs: 0 = top/left box, 1 = the other
// data.images: [{ png, field } | { png, page, rect: [x, y, w, h] }] – signatures and the sketch
// Returns the filled, flattened PDF as a Uint8Array.
(function () {
  var MULTILINE = ["Text9", "Text50", "Text93", "Text94", "Text95", "Text104", "Text112"];

  function widgetRect(field) {
    return field.acroField.getWidgets().map(function (w) { return w.getRectangle(); });
  }

  // The two boxes of a NEI/JA pair in reading order: top first, or left first on the same line.
  function choiceRects(field) {
    return widgetRect(field).sort(function (a, b) {
      return Math.abs(a.y - b.y) > 3 ? b.y - a.y : a.x - b.x;
    });
  }

  function pageOf(doc, field) {
    var ref = field.acroField.getWidgets()[0].P();
    var pages = doc.getPages();
    for (var i = 0; i < pages.length; i++) if (pages[i].ref === ref) return pages[i];
    return pages[0];
  }

  function drawCross(page, r, color) {
    var pad = Math.min(r.width, r.height) * 0.18;
    var opts = { thickness: 1.2, color: color };
    page.drawLine(Object.assign({ start: { x: r.x + pad, y: r.y + pad }, end: { x: r.x + r.width - pad, y: r.y + r.height - pad } }, opts));
    page.drawLine(Object.assign({ start: { x: r.x + pad, y: r.y + r.height - pad }, end: { x: r.x + r.width - pad, y: r.y + pad } }, opts));
  }

  function drawFitted(page, img, x, y, w, h) {
    var s = Math.min(w / img.width, h / img.height);
    var dw = img.width * s, dh = img.height * s;
    page.drawImage(img, { x: x + (w - dw) / 2, y: y + (h - dh) / 2, width: dw, height: dh });
  }

  window.fillSkademelding = function (templateBytes, data) {
    var L = window.PDFLib;
    var doc, form, font;
    return L.PDFDocument.load(templateBytes).then(function (d) {
      doc = d;
      form = doc.getForm();
      return doc.embedFont(L.StandardFonts.Helvetica);
    }).then(function (f) {
      font = f;
      Object.keys(data.text || {}).forEach(function (name) {
        // The form's font (Helvetica, WinAnsi) has no glyphs for emoji and most non-Latin letters.
        var value = String(data.text[name] || "").replace(/[^\n\x20-\x7E\xA0-\xFF–—‘’“”•€]/g, "").trim();
        if (!value) return;
        var field = form.getTextField(name);
        var max = field.getMaxLength();
        if (MULTILINE.indexOf(name) >= 0) field.enableMultiline();
        field.setText(max ? value.slice(0, max) : value);
      });
      (data.checks || []).forEach(function (name) { form.getCheckBox(name).check(); });

      // Pairs are drawn after flattening; their own on/off values are not reliable in this form.
      var crosses = Object.keys(data.choices || {}).map(function (name) {
        var field = form.getField(name);
        return { page: pageOf(doc, field), rect: choiceRects(field)[data.choices[name]] };
      });
      var images = (data.images || []).map(function (im) {
        if (!im.field) return im;
        var field = form.getField(im.field);
        var r = widgetRect(field)[0];
        return { png: im.png, page: doc.getPages().indexOf(pageOf(doc, field)), rect: [r.x, r.y, r.width, r.height] };
      });

      form.updateFieldAppearances(font);
      form.flatten();
      crosses.forEach(function (c) { if (c.rect) drawCross(c.page, c.rect, L.rgb(0, 0, 0)); });

      return Promise.all(images.map(function (im) { return doc.embedPng(im.png); })).then(function (embedded) {
        embedded.forEach(function (img, i) {
          var r = images[i].rect;
          drawFitted(doc.getPage(images[i].page), img, r[0] + 1, r[1] + 1, r[2] - 2, r[3] - 2);
        });
        return doc.save();
      });
    });
  };
})();
