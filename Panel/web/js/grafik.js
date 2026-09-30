// =====================================================================
// Grafik panel CKPN (Tahap 4) — SVG buatan sendiri, tanpa pustaka luar
// (panel berjalan offline dan CSP hanya mengizinkan file milik panel).
//
//   Grafik.garis(wadah, opsi)   grafik garis tren antarperiode + crosshair
//   Grafik.sankey(wadah, opsi)  sankey dua kolom (migrasi kualitas / net flow)
//
// Aturan tampilan (mengikuti pedoman visualisasi panel):
//   - warna seri diatur lewat kelas CSS (seri-1..3, aliran membaik/tetap/memburuk/keluar);
//   - garis 2px, penanda r=4 dengan cincin 2px warna permukaan, grid hairline;
//   - teks memakai warna tinta, bukan warna seri; label/tooltip memakai textContent;
//   - tooltip hanya pelengkap — semua angka juga ada di tabel di bawah grafik.
// =====================================================================
(function () {
  "use strict";
  var NS = "http://www.w3.org/2000/svg";

  function svgEl(tag, attr, kelas) {
    var n = document.createElementNS(NS, tag);
    if (attr) Object.keys(attr).forEach(function (k) { n.setAttribute(k, attr[k]); });
    if (kelas) n.setAttribute("class", kelas);
    return n;
  }
  function teksSvg(x, y, isi, kelas, anchor) {
    var t = svgEl("text", { x: x, y: y, "text-anchor": anchor || "start" }, kelas);
    t.textContent = isi;
    return t;
  }

  // ---------------- Tooltip bersama ----------------
  var tip = null;
  function tooltip() {
    if (tip) return tip;
    tip = document.createElement("div");
    tip.className = "grafik-tip";
    tip.setAttribute("role", "status");
    tip.hidden = true;
    document.body.appendChild(tip);
    return tip;
  }
  // baris: [{ kunci: kelasWarna|null, nilai: "Rp …", label: "Net Flow" }]
  function tampilTip(judul, baris, x, y) {
    var t = tooltip();
    t.innerHTML = "";
    var j = document.createElement("div");
    j.className = "grafik-tip-judul";
    j.textContent = judul;
    t.appendChild(j);
    baris.forEach(function (b) {
      var r = document.createElement("div");
      r.className = "grafik-tip-baris";
      if (b.kunci) {
        var k = document.createElement("span");
        k.className = "kunci-garis " + b.kunci;
        r.appendChild(k);
      }
      var v = document.createElement("b");
      v.textContent = b.nilai;
      r.appendChild(v);
      if (b.label) {
        var l = document.createElement("span");
        l.textContent = b.label;
        r.appendChild(l);
      }
      t.appendChild(r);
    });
    t.hidden = false;
    var lebar = t.offsetWidth, tinggi = t.offsetHeight;
    var kiri = Math.min(Math.max(8, x + 12), window.innerWidth - lebar - 8);
    var atas = y + tinggi + 16 > window.innerHeight + window.scrollY ? y - tinggi - 12 : y + 12;
    t.style.left = kiri + "px";
    t.style.top = atas + "px";
  }
  function sembunyiTip() { if (tip) tip.hidden = true; }

  // ---------------- Skala ----------------
  function langkahBulat(maks, n) {
    if (maks <= 0) return 1;
    var kasar = maks / n, p = Math.pow(10, Math.floor(Math.log10(kasar))), r = kasar / p;
    return (r <= 1 ? 1 : r <= 2 ? 2 : r <= 2.5 ? 2.5 : r <= 5 ? 5 : 10) * p;
  }

  // =====================================================================
  // Grafik garis
  // opsi: { x: [label], seri: [{ nama, singkat, kelas, nilai: [angka|null] }],
  //         formatSumbu: fn, formatNilai: fn, tinggi }
  // =====================================================================
  function garis(wadah, o) {
    wadah.innerHTML = "";
    var W = Math.max(280, wadah.clientWidth || 360), H = o.tinggi || 220;
    var m = { l: 46, r: 44, t: 12, b: 26 };
    var pw = W - m.l - m.r, ph = H - m.t - m.b, n = o.x.length;

    // Domain Y: grafik garis tidak wajib mulai dari nol — rentang dipersempit agar tren terbaca,
    // tetap dibulatkan ke angka bersih dan dicantumkan di sumbu.
    var maks = -Infinity, min = Infinity;
    o.seri.forEach(function (s) { s.nilai.forEach(function (v) { if (v !== null && v !== undefined) { if (v > maks) maks = v; if (v < min) min = v; } }); });
    if (!isFinite(maks)) { maks = 1; min = 0; }
    var rentang = Math.max(maks - min, Math.abs(maks) * 0.1, 1);
    var step = langkahBulat(rentang * 1.3, 4);
    var bawah = Math.max(0, Math.floor((min - rentang * 0.15) / step) * step);
    var atas = Math.ceil((maks + rentang * 0.1) / step) * step;
    if (atas <= bawah) atas = bawah + step;
    function sx(i) { return m.l + (n <= 1 ? pw / 2 : (i * pw) / (n - 1)); }
    function sy(v) { return m.t + ph - ((v - bawah) / (atas - bawah)) * ph; }

    var svg = svgEl("svg", { width: W, height: H, viewBox: "0 0 " + W + " " + H, role: "img" }, "grafik-svg");
    svg.setAttribute("aria-label", o.label || "Grafik tren");

    // grid & sumbu Y
    for (var g = bawah; g <= atas + step / 2; g += step) {
      var y = sy(g);
      svg.appendChild(svgEl("line", { x1: m.l, x2: W - m.r, y1: y, y2: y }, g === bawah ? "sumbu" : "grid"));
      svg.appendChild(teksSvg(m.l - 6, y + 4, o.formatSumbu(g, step), "label-sumbu", "end"));
    }
    // sumbu X — label dijarangkan agar tidak bertumpuk (±52px per label)
    var tiap = Math.max(1, Math.ceil(n / Math.max(1, Math.floor(pw / 52))));
    var tampilX = [];
    for (var i = 0; i < n; i += tiap) tampilX.push(i);
    if (tampilX[tampilX.length - 1] !== n - 1) {
      // periode terakhir selalu diberi label; label sebelumnya dibuang bila terlalu dekat
      if (n - 1 - tampilX[tampilX.length - 1] < tiap * 0.6 && tampilX.length > 1) tampilX.pop();
      tampilX.push(n - 1);
    }
    tampilX.forEach(function (i) {
      svg.appendChild(teksSvg(sx(i), H - 8, o.x[i], "label-sumbu",
        n === 1 ? "middle" : i === 0 ? "start" : i === n - 1 ? "end" : "middle"));
    });

    // garis & penanda
    var akhir = [];
    o.seri.forEach(function (s) {
      var d = "", naik = false;
      s.nilai.forEach(function (v, i) {
        if (v === null || v === undefined) { naik = false; return; }
        d += (naik ? "L" : "M") + sx(i).toFixed(1) + "," + sy(v).toFixed(1);
        naik = true;
      });
      if (d) svg.appendChild(svgEl("path", { d: d }, "garis " + s.kelas));
      // penanda di titik terakhir yang ada (dan di semua titik bila hanya sedikit)
      var terakhir = -1;
      s.nilai.forEach(function (v, i) { if (v !== null && v !== undefined) terakhir = i; });
      s.nilai.forEach(function (v, i) {
        if (v === null || v === undefined) return;
        if (i === terakhir || n <= 6)
          svg.appendChild(svgEl("circle", { cx: sx(i), cy: sy(v), r: 4 }, "penanda " + s.kelas));
      });
      if (terakhir >= 0) akhir.push({ x: sx(terakhir), y: sy(s.nilai[terakhir]), teks: s.singkat || s.nama });
    });

    // label ujung langsung — hanya bila tidak bertabrakan (jarak ≥ 12px); legenda tetap ada
    akhir.sort(function (a, b) { return a.y - b.y; });
    var bentrok = akhir.some(function (a, i) { return i > 0 && a.y - akhir[i - 1].y < 12; });
    if (!bentrok) akhir.forEach(function (a) { svg.appendChild(teksSvg(a.x + 8, a.y + 4, a.teks, "label-ujung")); });

    // crosshair + tooltip; area sentuh = seluruh plot
    var garisX = svgEl("line", { y1: m.t, y2: m.t + ph, x1: 0, x2: 0, visibility: "hidden" }, "crosshair");
    svg.appendChild(garisX);
    var lapis = svgEl("rect", { x: m.l - 8, y: m.t, width: pw + 16, height: ph, fill: "#fff", "fill-opacity": 0, "pointer-events": "all", tabindex: 0 }, "lapis-sentuh");
    lapis.setAttribute("aria-label", "Geser atau pakai panah kiri/kanan untuk membaca nilai tiap periode");
    svg.appendChild(lapis);
    var aktif = n - 1;
    function tunjuk(i, px, py) {
      aktif = Math.max(0, Math.min(n - 1, i));
      garisX.setAttribute("x1", sx(aktif)); garisX.setAttribute("x2", sx(aktif)); garisX.setAttribute("visibility", "visible");
      tampilTip(o.x[aktif], o.seri.map(function (s) {
        var v = s.nilai[aktif];
        return { kunci: s.kelas, nilai: v === null || v === undefined ? "—" : o.formatNilai(v), label: s.nama };
      }), px, py);
    }
    lapis.addEventListener("pointermove", function (e) {
      var r = svg.getBoundingClientRect();
      var i = n <= 1 ? 0 : Math.round(((e.clientX - r.left - m.l) / pw) * (n - 1));
      tunjuk(i, e.pageX, e.pageY);
    });
    lapis.addEventListener("pointerleave", function () { garisX.setAttribute("visibility", "hidden"); sembunyiTip(); });
    lapis.addEventListener("focus", function () { var r = svg.getBoundingClientRect(); tunjuk(aktif, r.left + window.scrollX + sx(aktif), r.top + window.scrollY + m.t); });
    lapis.addEventListener("blur", sembunyiTip);
    lapis.addEventListener("keydown", function (e) {
      if (e.key !== "ArrowLeft" && e.key !== "ArrowRight") return;
      e.preventDefault();
      var r = svg.getBoundingClientRect();
      tunjuk(aktif + (e.key === "ArrowRight" ? 1 : -1), r.left + window.scrollX + sx(aktif), r.top + window.scrollY + m.t);
    });
    wadah.appendChild(svg);
  }

  // =====================================================================
  // Sankey dua kolom
  // opsi: { kiri: [{ id, label }], kanan: [{ id, label }],
  //         aliran: [{ dari, ke, nilai, kelas, info }],   // info: teks tambahan untuk tooltip
  //         formatNilai: fn, tinggi, label }
  // Tinggi simpul sebanding nilai; simpul kiri = total aliran keluar, kanan = total masuk.
  // =====================================================================
  function sankey(wadah, o) {
    wadah.innerHTML = "";
    var W = Math.max(300, wadah.clientWidth || 360), H = o.tinggi || 300;
    var lebarLabelKiri = o.lebarLabelKiri || 64, lebarLabelKanan = o.lebarLabelKanan || 82, nw = 8, jarak = 6;
    var xKiri = lebarLabelKiri, xKanan = W - lebarLabelKanan - nw;

    var nilaiKiri = {}, nilaiKanan = {}, total = 0;
    o.aliran.forEach(function (a) {
      if (!(a.nilai > 0)) return;
      nilaiKiri[a.dari] = (nilaiKiri[a.dari] || 0) + a.nilai;
      nilaiKanan[a.ke] = (nilaiKanan[a.ke] || 0) + a.nilai;
      total += a.nilai;
    });
    var svg = svgEl("svg", { width: W, height: H, viewBox: "0 0 " + W + " " + H, role: "img" }, "grafik-svg");
    svg.setAttribute("aria-label", o.label || "Sankey");
    if (total <= 0) {
      svg.appendChild(teksSvg(W / 2, H / 2, "Tidak ada saldo untuk digambar.", "label-sumbu", "middle"));
      wadah.appendChild(svg);
      return;
    }

    function tata(simpul, nilai) {
      var skala = (H - jarak * (simpul.length - 1)) / total, y = 0, pos = {};
      simpul.forEach(function (s) {
        var h = (nilai[s.id] || 0) * skala;
        pos[s.id] = { y: y, h: h, isi: y };
        y += h + jarak;
      });
      return { pos: pos, skala: skala };
    }
    var tk = tata(o.kiri, nilaiKiri), tn = tata(o.kanan, nilaiKanan), skala = tk.skala;
    var urutKiri = {}, urutKanan = {};
    o.kiri.forEach(function (s, i) { urutKiri[s.id] = i; });
    o.kanan.forEach(function (s, i) { urutKanan[s.id] = i; });

    // urutan pita: di simpul kiri menurut urutan tujuan, di simpul kanan menurut urutan asal
    var aliran = o.aliran.filter(function (a) { return a.nilai > 0; });
    aliran.sort(function (a, b) { return urutKiri[a.dari] - urutKiri[b.dari] || urutKanan[a.ke] - urutKanan[b.ke]; });
    aliran.forEach(function (a) { var p = tk.pos[a.dari]; a.y0 = p.isi; p.isi += a.nilai * skala; });
    aliran.slice().sort(function (a, b) { return urutKanan[a.ke] - urutKanan[b.ke] || urutKiri[a.dari] - urutKiri[b.dari]; })
      .forEach(function (a) { var p = tn.pos[a.ke]; a.y1 = p.isi; p.isi += a.nilai * skala; });

    var gPita = svgEl("g", null, "pita-grup");
    svg.appendChild(gPita);
    var x0 = xKiri + nw, x1 = xKanan, xm = (x0 + x1) / 2;
    var labelKiri = {}, labelKanan = {};
    o.kiri.forEach(function (s) { labelKiri[s.id] = s.label; });
    o.kanan.forEach(function (s) { labelKanan[s.id] = s.label; });

    aliran.forEach(function (a) {
      var h = Math.max(a.nilai * skala, 1);   // pita sangat kecil tetap terlihat 1px
      var ya = a.y0, yb = a.y1;
      var d = "M" + x0 + "," + ya + "C" + xm + "," + ya + " " + xm + "," + yb + " " + x1 + "," + yb +
              "L" + x1 + "," + (yb + h) + "C" + xm + "," + (yb + h) + " " + xm + "," + (ya + h) + " " + x0 + "," + (ya + h) + "Z";
      var p = svgEl("path", { d: d, tabindex: 0 }, "pita " + a.kelas);
      var judul = labelKiri[a.dari] + " → " + labelKanan[a.ke];
      p.setAttribute("aria-label", judul + ": " + o.formatNilai(a.nilai) + (a.info ? ", " + a.info : ""));
      function tampil(ev) {
        gPita.classList.add("ada-sorot");
        p.classList.add("sorot");
        var baris = [{ kunci: a.kelas, nilai: o.formatNilai(a.nilai), label: a.namaKelas || "" }];
        if (a.info) baris.push({ nilai: a.info });
        var r = p.getBoundingClientRect();
        tampilTip(judul, baris, ev && ev.pageX ? ev.pageX : r.left + window.scrollX + r.width / 2,
                               ev && ev.pageY ? ev.pageY : r.top + window.scrollY);
      }
      function lepas() { gPita.classList.remove("ada-sorot"); p.classList.remove("sorot"); sembunyiTip(); }
      p.addEventListener("pointermove", tampil);
      p.addEventListener("pointerleave", lepas);
      p.addEventListener("focus", function () { tampil(null); });
      p.addEventListener("blur", lepas);
      gPita.appendChild(p);
    });

    // simpul & label
    function gambarSimpul(simpul, tata, x, kiri, nilai) {
      var labelTerakhir = -Infinity;
      simpul.forEach(function (s) {
        var p = tata.pos[s.id];
        if (p.h > 0) svg.appendChild(svgEl("rect", { x: x, y: p.y, width: nw, height: Math.max(p.h, 1), rx: 2 }, "simpul " + (s.kelas || "")));
        var ty = p.y + Math.max(p.h, 0) / 2 + 4;
        // label yang bertumpuk dengan label di atasnya tidak digambar — nilainya tetap ada di tooltip & tabel
        if (ty - labelTerakhir < 12) return;
        labelTerakhir = ty;
        var t = teksSvg(kiri ? x - 6 : x + nw + 6, ty, s.label, "label-simpul" + (p.h > 0 ? "" : " kosong"), kiri ? "end" : "start");
        svg.appendChild(t);
        if (o.subLabel && p.h >= 22) {
          svg.appendChild(teksSvg(kiri ? x - 6 : x + nw + 6, ty + 12, o.subLabel(s.id, nilai[s.id] || 0, kiri), "label-sumbu", kiri ? "end" : "start"));
        }
      });
    }
    gambarSimpul(o.kiri, tk, xKiri, true, nilaiKiri);
    gambarSimpul(o.kanan, tn, xKanan, false, nilaiKanan);
    wadah.appendChild(svg);
  }

  window.Grafik = { garis: garis, sankey: sankey, sembunyiTip: sembunyiTip };
})();
