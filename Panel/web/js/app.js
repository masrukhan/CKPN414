// =====================================================================
// Panel CKPN — logika sisi panel
// Tab: Overview · Manual (jalankan langkah pilihan) · Per grup (alur periode) · Riwayat · Analisis
// Kartu progres/hasil/simpan berada di area proses bersama, sehingga terlihat dari tab mana pun.
//
// Komunikasi dengan C# (PanelBridge.cs):
//   kirim  : chrome.webview.postMessage({ id, cmd, args })
//   terima : { type: "reply", id, ok, data | error }
//            { type: "event", name, data }
//
// Event pipeline (CKPNPipeline.cs):
//   runMulai → (langkahMulai → progresDetail* → pesanModul* → langkahSelesai)* → runSelesai
// =====================================================================
(function () {
  "use strict";

  var webview = window.chrome && window.chrome.webview;
  var nextId = 1;
  var menunggu = {};
  var pendengarEvent = {};

  // ---------------- Jembatan ----------------
  function panggil(cmd, args, batasMs) {
    return new Promise(function (resolve, reject) {
      if (!webview) { reject(new Error("Panel tidak berjalan di dalam Excel.")); return; }
      var id = nextId++;
      var timer = setTimeout(function () {
        delete menunggu[id];
        reject(new Error("Tidak ada jawaban dari library untuk '" + cmd + "'."));
      }, batasMs || 60000);
      menunggu[id] = { resolve: resolve, reject: reject, timer: timer };
      webview.postMessage({ id: id, cmd: cmd, args: args || {} });
    });
  }
  function onEvent(nama, fn) { (pendengarEvent[nama] = pendengarEvent[nama] || []).push(fn); }

  if (webview) {
    webview.addEventListener("message", function (e) {
      var m = e.data;
      if (!m || typeof m !== "object") return;
      if (m.type === "reply") {
        var p = menunggu[m.id];
        if (!p) return;
        clearTimeout(p.timer);
        delete menunggu[m.id];
        if (m.ok) p.resolve(m.data); else p.reject(new Error(m.error || "Gagal"));
      } else if (m.type === "event") {
        (pendengarEvent[m.name] || []).forEach(function (fn) { fn(m.data || {}); });
      }
    });
  }
  window.CKPN = { panggil: panggil, onEvent: onEvent };

  // ---------------- Util ----------------
  function $(id) { return document.getElementById(id); }
  function teks(id, v) { $(id).textContent = (v === undefined || v === null || v === "") ? "—" : v; }
  function el(tag, cls, isi) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (isi !== undefined) n.textContent = isi;
    return n;
  }
  function detik(ms) {
    var s = Math.round(ms / 1000);
    return s < 60 ? s + " dtk" : Math.floor(s / 60) + " mnt " + (s % 60) + " dtk";
  }
  function tampil(id, ya) { $(id).hidden = !ya; }

  var IKON = {
    menunggu: '<svg class="ikon" viewBox="0 0 20 20" aria-hidden="true"><circle cx="10" cy="10" r="8" fill="none" stroke="#C9C6BC" stroke-width="1.5" stroke-dasharray="3 3"/></svg>',
    berjalan: '<svg class="ikon" viewBox="0 0 20 20" aria-hidden="true"><circle cx="10" cy="10" r="8" fill="none" stroke="#CFE0DC" stroke-width="2.5"/><path class="ikon-putar" d="M10 2a8 8 0 0 1 8 8" fill="none" stroke="#0E5A50" stroke-width="2.5" stroke-linecap="round"/></svg>',
    selesai:  '<svg class="ikon" viewBox="0 0 20 20" aria-hidden="true"><circle cx="10" cy="10" r="9" fill="#1F6E43"/><path d="M6 10.2l2.7 2.7 5.3-5.6" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    gagal:    '<svg class="ikon" viewBox="0 0 20 20" aria-hidden="true"><circle cx="10" cy="10" r="9" fill="#A1321F"/><path d="M7 7l6 6M13 7l-6 6" stroke="#fff" stroke-width="2" stroke-linecap="round"/></svg>'
  };

  // ---------------- Tab ----------------
  var tombolTab = document.querySelectorAll('.tab [role="tab"]');
  tombolTab.forEach(function (b) {
    b.addEventListener("click", function () {
      var nama = b.getAttribute("data-tab");
      if (nama === "periode") { muatPeriode(); muatPenyesuaian(); if ($("db-kelola").open) muatInfoDatabase(); }
      if (nama === "overview" && !dataOverview) muatOverview();
      tombolTab.forEach(function (x) {
        var aktif = x.getAttribute("data-tab") === nama;
        x.setAttribute("aria-selected", aktif ? "true" : "false");
        $("tab-" + x.getAttribute("data-tab")).hidden = !aktif;
      });
    });
  });

  // =====================================================================
  // 1. Persiapan: baca Master & validasi setiap langkah
  // =====================================================================
  var dataLangkah = [];     // hasil siapkanRun

  function muatPersiapan() {
    teks("periode", "Memeriksa…");
    $("btn-jalankan").disabled = true;
    return panggil("siapkanRun").then(function (d) {
      teks("periode", d.periode || "Master!C4 belum diisi");
      teks("wb-nama", d.namaFile);
      var chip = $("kc-chip");
      chip.innerHTML = "";
      (d.kc || []).forEach(function (k) { chip.appendChild(el("span", "chip", k)); });
      if (!d.kc || d.kc.length === 0) chip.appendChild(el("span", "teks-kecil", "Belum ada KC yang dicentang di Master."));

      dataLangkah = d.langkah || [];
      gambarDaftarLangkah();
      if (d.berjalan) tampilkanModeBerjalan();
    }).catch(function (err) {
      teks("periode", err.message);
      dataLangkah = [];
      gambarDaftarLangkah();
    });
  }

  function gambarDaftarLangkah() {
    var ul = $("daftar-langkah");
    ul.innerHTML = "";
    dataLangkah.forEach(function (l) {
      var li = el("li", l.siap ? "" : "tidak-siap");
      var label = el("label");
      var cek = el("input");
      cek.type = "checkbox";
      cek.value = l.id;
      cek.checked = l.siap;
      cek.disabled = !l.siap;
      cek.addEventListener("change", perbaruiTombolJalankan);
      var isi = el("div");
      isi.appendChild(el("div", "nama", l.nama));
      if (l.ringkas) isi.appendChild(el("div", "ringkas", l.ringkas));
      (l.error || []).forEach(function (e) { isi.appendChild(el("div", "masalah", e)); });
      label.appendChild(cek);
      label.appendChild(isi);
      li.appendChild(label);
      ul.appendChild(li);
    });
    perbaruiTombolJalankan();
  }

  function langkahDipilih() {
    var ids = [];
    document.querySelectorAll('#daftar-langkah input[type="checkbox"]').forEach(function (c) {
      if (c.checked && !c.disabled) ids.push(c.value);
    });
    return ids;
  }

  function perbaruiTombolJalankan() {
    var n = langkahDipilih().length;
    var b = $("btn-jalankan");
    b.disabled = n === 0;
    b.textContent = n === 0 ? "Pilih minimal satu langkah" :
                    n === dataLangkah.length ? "Jalankan semua (" + n + " langkah)" :
                    "Jalankan " + n + " langkah terpilih";
  }

  $("cek-semua").addEventListener("change", function () {
    var v = this.checked;
    document.querySelectorAll('#daftar-langkah input[type="checkbox"]').forEach(function (c) {
      if (!c.disabled) c.checked = v;
    });
    perbaruiTombolJalankan();
  });

  // =====================================================================
  // 2. Konfirmasi
  // =====================================================================
  $("btn-jalankan").addEventListener("click", function () {
    var ids = langkahDipilih();
    var ul = $("daftar-sheet");
    ul.innerHTML = "";
    dataLangkah.forEach(function (l) {
      if (ids.indexOf(l.id) >= 0) ul.appendChild(el("li", "", l.sheet + "  (" + l.nama + ")"));
    });
    tampil("kartu-langkah", false);
    tampil("kartu-konfirmasi", true);
    $("btn-ya").focus();
  });

  $("btn-tidak").addEventListener("click", function () {
    tampil("kartu-konfirmasi", false);
    tampil("kartu-langkah", true);
  });

  $("btn-ya").addEventListener("click", function () {
    var ids = langkahDipilih();
    $("btn-ya").disabled = true;
    panggil("jalankan", { langkah: ids, terapkanPenyesuaian: $("cek-penyesuaian").checked }).then(function () {
      // Tampilan progres dibangun oleh event runMulai
    }).catch(function (err) {
      tampil("kartu-konfirmasi", false);
      tampilkanHasil({ status: "gagal", error: err.message, langkah: [], pesan: [] });
    }).then(function () { $("btn-ya").disabled = false; });
  });

  // =====================================================================
  // 3. Progres (event dari CKPNPipeline)
  // =====================================================================
  var run = null;   // { langkah: [{id,nama}], selesai: n, mulai: Date }

  function tampilkanModeBerjalan() {
    tampil("kartu-langkah", false);
    tampil("kartu-konfirmasi", false);
    tampil("kartu-hasil", false);
    tampil("kartu-progres", true);
  }

  onEvent("runMulai", function (d) {
    run = { langkah: d.langkah || [], selesai: 0, mulai: Date.now() };
    var ol = $("progres-langkah");
    ol.innerHTML = "";
    run.langkah.forEach(function (l) {
      var li = el("li", "menunggu");
      li.id = "pl-" + l.id;
      li.innerHTML = IKON.menunggu;
      var isi = el("div", "isi");
      var baris = el("div", "baris-nama");
      baris.appendChild(el("span", "nama", l.nama));
      baris.appendChild(el("span", "durasi", ""));
      isi.appendChild(baris);
      isi.appendChild(el("div", "detail-langkah", ""));
      li.appendChild(isi);
      ol.appendChild(li);
    });
    teks("progres-judul", "Sedang menghitung…");
    teks("progres-sub", (d.periode || "") + " · " + (d.kc || ""));
    setPersen(0);
    $("btn-batal").disabled = false;
    $("btn-batal").textContent = "Batalkan setelah langkah ini";
    tampilkanModeBerjalan();
  });

  function setLangkah(id, status, detail, durasi) {
    var li = $("pl-" + id);
    if (!li) return;
    li.className = status;
    li.querySelector("svg").outerHTML = IKON[status];
    if (detail !== undefined) li.querySelector(".detail-langkah").textContent = detail;
    if (durasi !== undefined) li.querySelector(".durasi").textContent = durasi;
  }

  function setPersen(p) {
    p = Math.max(0, Math.min(100, Math.round(p)));
    $("progres-bar").style.width = p + "%";
    teks("progres-persen", p + "%");
    document.querySelector(".bar").setAttribute("aria-valuenow", String(p));
  }

  onEvent("langkahMulai", function (d) {
    if (!run) return;
    setLangkah(d.id, "berjalan", "Berjalan…");
    teks("progres-judul", d.nama);
    teks("progres-sub", "Langkah " + d.ke + " dari " + d.dari);
  });

  // Progres di dalam modul (mis. bulan ke-9 dari 12). Persentase keseluruhan
  // = langkah yang sudah selesai + pecahan langkah berjalan.
  onEvent("progresDetail", function (d) {
    if (!run) return;
    var aktif = document.querySelector(".progres-langkah li.berjalan");
    if (aktif) aktif.querySelector(".detail-langkah").textContent =
      d.label ? (d.ke + " dari " + d.dari + " · " + d.label) : (d.ke + " dari " + d.dari);
    var n = run.langkah.length || 1;
    setPersen(((run.selesai + (d.ke - 1) / Math.max(1, d.dari)) / n) * 100);
  });

  onEvent("langkahSelesai", function (d) {
    if (!run) return;
    run.selesai++;
    setLangkah(d.id, d.ok ? "selesai" : "gagal", d.ok ? "Selesai" : d.error, detik(d.durasiMs));
    setPersen((run.selesai / (run.langkah.length || 1)) * 100);
  });

  onEvent("runSelesai", function (d) {
    tampilkanHasil(d);
    if (d.status === "selesai") muatReview();
    run = null;
    muatPersiapan();   // segarkan status langkah untuk run berikutnya
  });

  $("btn-batal").addEventListener("click", function () {
    $("btn-batal").disabled = true;
    $("btn-batal").textContent = "Menunggu langkah berjalan selesai…";
    panggil("batal").catch(function () {});
  });

  // =====================================================================
  // 4. Hasil
  // =====================================================================
  function tampilkanHasil(d) {
    var st = $("hasil-status");
    if (d.status === "selesai") {
      st.className = "status-hasil ok";
      st.textContent = "Perhitungan selesai · " + detik(d.durasiMs || 0);
    } else if (d.status === "dibatalkan") {
      st.className = "status-hasil batal";
      st.textContent = "Dibatalkan · " + (d.langkah || []).length + " langkah sempat selesai";
    } else {
      st.className = "status-hasil err";
      st.textContent = "Gagal: " + (d.error || "tidak diketahui");
    }

    var wadah = $("hasil-pesan");
    wadah.innerHTML = "";
    (d.pesan || []).forEach(function (p) {
      var box = el("div", "pesan-modul" + (p.peringatan ? " peringatan" : ""));
      box.appendChild(el("div", "judul-pesan", p.judul));
      box.appendChild(el("pre", "", p.pesan));
      wadah.appendChild(box);
    });

    $("hasil-review").innerHTML = "";
    tampil("btn-ke-simpan", d.status === "selesai");
    tampil("kartu-progres", false);
    tampil("kartu-konfirmasi", false);
    tampil("kartu-simpan", false);
    tampil("kartu-hasil", true);
    tampil("kartu-langkah", true);   // hasil tampil di area proses; daftar langkah manual kembali tersedia
  }

  $("btn-tutup-hasil").addEventListener("click", function () {
    tampil("kartu-hasil", false);
    tampil("kartu-langkah", true);
  });

  $("btn-periksa").addEventListener("click", muatPersiapan);

  // =====================================================================
  // 5. Review hasil hitung (Tahap 3a)
  // =====================================================================
  var fmt = new Intl.NumberFormat("id-ID", { maximumFractionDigits: 0 });
  function rp(v) { return v === null || v === undefined ? "—" : fmt.format(v); }

  function tombolLompat(label, sheet, sel) {
    var b = el("button", "", label);
    b.type = "button";
    b.title = "Buka " + sheet + " sel " + sel;
    b.addEventListener("click", function () {
      panggil("lompatKe", { sheet: sheet, sel: sel }).catch(function (e) { alert(e.message); });
    });
    return b;
  }

  function muatReview() {
    panggil("reviewInfo").then(gambarReview).catch(function () {});
  }

  function gambarReview(d) {
    var wadah = $("hasil-review");
    wadah.innerHTML = "";
    var ind = d && d.individu, cs = d && d.lgdcs;
    if (!ind && !cs) return;

    if (ind) {
      var b1 = el("div", "blok-review");
      b1.appendChild(el("h4", "", "CKPN Individu — cek kolom I (agunan) & J (biaya penjualan)"));
      var a1 = el("div", "angka-review");
      a1.appendChild(el("span", "", ind.jumlahKontrak + " kontrak"));
      a1.appendChild(el("span", "", ind.diterapkan + " dari penyesuaian"));
      if (ind.berubah.length) a1.appendChild(el("span", "perhatian", ind.berubah.length + " agunan sistem berubah"));
      if (ind.baru.length) a1.appendChild(el("span", "perhatian", ind.baru.length + " kontrak baru"));
      if (!ind.penyesuaianDiterapkan) a1.appendChild(el("span", "perhatian", "tanpa penyesuaian"));
      b1.appendChild(a1);
      var ul1 = el("ul", "daftar-klik");
      ind.berubah.forEach(function (x) {
        var li = el("li");
        li.appendChild(tombolLompat(x.kunci, "A. CKPN - INDV", "I" + x.baris));
        li.appendChild(el("span", "", "sistem " + rp(x.sistemLama) + " → " + rp(x.sistemBaru) + " · dipakai " + rp(x.dipakai)));
        ul1.appendChild(li);
      });
      ind.baru.forEach(function (x) {
        var li = el("li");
        li.appendChild(tombolLompat(x.kunci, "A. CKPN - INDV", "I" + x.baris));
        li.appendChild(el("span", "", "baru masuk daftar individu"));
        ul1.appendChild(li);
      });
      b1.appendChild(ul1);
      if (ind.masalah) b1.appendChild(el("div", "peringatan-box", ind.masalah));
      wadah.appendChild(b1);
    }

    if (cs) {
      var b2 = el("div", "blok-review");
      b2.appendChild(el("h4", "", "LGD Collateral Shortfall — cek ke remedial"));
      var a2 = el("div", "angka-review");
      a2.appendChild(el("span", "", cs.jumlahSistem + " debitur dari sistem"));
      if (cs.dikecualikan) a2.appendChild(el("span", "", cs.dikecualikan + " dikecualikan otomatis"));
      if (cs.diubah) a2.appendChild(el("span", "", cs.diubah + " nilai dari penyesuaian"));
      if (cs.ditambah) a2.appendChild(el("span", "", cs.ditambah + " baris manual"));
      if (cs.baru.length) a2.appendChild(el("span", "perhatian", cs.baru.length + " debitur baru"));
      if (!cs.penyesuaianDiterapkan) a2.appendChild(el("span", "perhatian", "tanpa penyesuaian"));
      b2.appendChild(a2);
      var ul2 = el("ul", "daftar-klik");
      cs.baru.forEach(function (x) {
        var li = el("li");
        li.appendChild(tombolLompat(x.kunci, "B4.LGD-CS MACET", "B5"));
        li.appendChild(el("span", "", x.nama || ""));
        ul2.appendChild(li);
      });
      b2.appendChild(ul2);
      if (cs.masalah) b2.appendChild(el("div", "peringatan-box", cs.masalah));
      wadah.appendChild(b2);
    }
    wadah.appendChild(el("p", "teks-kecil catatan",
      "Edit langsung di sheet seperti biasa: ubah I/J di Individu; di LGD CS ubah nilai agunan (D) atau realisasi (G), " +
      "hapus baris yang bukan eksekusi agunan, atau tambah baris manual di bawah data (sebelum baris Total). " +
      "Setelah selesai, tekan tombol simpan di bawah."));
  }

  // =====================================================================
  // 6. Simpan grup ke staging
  // =====================================================================
  var ALASAN_UMUM = [
    "Lunas bukan dari eksekusi agunan",
    "Restrukturisasi / kembali lancar",
    "Hapus buku tanpa eksekusi agunan",
    "Data remedial belum lengkap"
  ];
  var dataSimpan = null;

  function bukaSimpan() {
    tampil("kartu-hasil", false);
    tampil("kartu-langkah", false);
    tampil("kartu-simpan", true);
    $("simpan-hasil").innerHTML = "";
    $("kartu-simpan").scrollIntoView({ block: "start" });
    periksaSimpan();
  }

  function periksaSimpan() {
    $("btn-simpan").disabled = true;
    teks("simpan-info", "Memeriksa sheet…");
    $("simpan-peringatan").innerHTML = "";
    $("simpan-perubahan").innerHTML = "";
    panggil("periksaSimpan").then(function (d) {
      dataSimpan = d;
      teks("simpan-info", d.periode + " · " + (d.namaGrup ? d.namaGrup + " (" + d.kodeKC + ")" : d.kodeKC) +
        " · akan disimpan sebagai versi " + d.versi +
        " · " + d.jumlahIndividu + " kontrak individu, " + d.jumlahLgdCs + " baris LGD CS");
      var pw = $("simpan-peringatan");
      (d.peringatan || []).forEach(function (t) { pw.appendChild(el("div", "peringatan-box", t)); });
      if (!d.bolehMenulis) pw.appendChild(el("div", "peringatan-box", d.infoPengirim));
      else if (d.infoPengirim) pw.appendChild(el("div", "teks-kecil", d.infoPengirim));
      gambarPerubahan(d.perubahan || []);
      $("btn-simpan").disabled = !d.bolehMenulis || d.berjalan;
    }).catch(function (e) {
      teks("simpan-info", e.message);
    });
  }

  var LABEL_AKSI = {
    "baru": "penyesuaian baru", "ubah": "diubah", "kembali": "kembali ke sistem",
    "hapus-baris": "dihapus dari LGD CS", "hapus-manual": "baris manual dihapus"
  };

  function gambarPerubahan(list) {
    var w = $("simpan-perubahan");
    w.innerHTML = "";
    if (list.length === 0) {
      w.appendChild(el("p", "teks-kecil", "Tidak ada perubahan penyesuaian — hasil tetap disimpan sebagai versi baru."));
      return;
    }
    [["individu", "CKPN Individu"], ["lgdcs", "LGD Collateral Shortfall"]].forEach(function (m) {
      var item = list.filter(function (x) { return x.modul === m[0]; });
      if (!item.length) return;
      var g = el("div", "perubahan-grup");
      g.appendChild(el("h4", "", m[1] + " (" + item.length + ")"));
      item.forEach(function (x) {
        var box = el("div", "perubahan-item");
        var head = el("div");
        head.appendChild(el("span", "kunci", x.kunci));
        head.appendChild(el("span", "aksi " + x.aksi, LABEL_AKSI[x.aksi] || x.aksi));
        box.appendChild(head);
        box.appendChild(el("div", "teks-kecil", x.sebelum + " → " + x.sesudah));
        if (x.bisaAlasan) {
          box.appendChild(el("div", "teks-kecil", "Disimpan sebagai pengecualian: rekening ini tidak ikut LGD CS bulan berikutnya. " +
            "Untuk memasukkannya lagi, hapus pengecualiannya di tab Periode › Penyesuaian tersimpan."));
          var sel = el("select", "isian");
          sel.setAttribute("aria-label", "Alasan penghapusan " + x.kunci + " (opsional)");
          sel.dataset.kunci = x.kunci;
          sel.className = "isian alasan-hapus";
          var kosong = el("option", "", "— alasan (opsional) —"); kosong.value = "";
          sel.appendChild(kosong);
          ALASAN_UMUM.forEach(function (a) { var o = el("option", "", a); o.value = a; sel.appendChild(o); });
          var lain = el("option", "", "Lainnya…"); lain.value = "__lain"; sel.appendChild(lain);
          var input = el("input", "isian");
          input.type = "text";
          input.placeholder = "Tulis alasan";
          input.hidden = true;
          sel.addEventListener("change", function () { input.hidden = sel.value !== "__lain"; });
          box.appendChild(sel);
          box.appendChild(input);
        }
        g.appendChild(box);
      });
      w.appendChild(g);
    });
  }

  // Alasan hapus LGD CS bersifat opsional (Tahap 3c): yang menyaring perhitungan
  // bulan berikutnya adalah pengecualiannya; alasan hanya jejak audit.
  function kumpulkanAlasan() {
    var hasil = {};
    document.querySelectorAll("select.alasan-hapus").forEach(function (sel) {
      var v = sel.value === "__lain" ? sel.nextSibling.value.trim() : sel.value;
      if (v) hasil[sel.dataset.kunci] = v;
    });
    return { alasan: hasil };
  }

  $("btn-simpan").addEventListener("click", function () {
    var a = kumpulkanAlasan();
    var btn = $("btn-simpan");
    btn.disabled = true;
    btn.textContent = "Menyimpan…";
    panggil("simpanGrup", {
      alasan: a.alasan,
      catatan: $("simpan-catatan").value,
      refreshSummary: $("cek-refresh").checked
    }, 600000).then(function (r) {
      var w = $("simpan-hasil");
      w.innerHTML = "";
      var box = el("div", "hasil-simpan");
      box.appendChild(el("div", "status-hasil ok", "Tersimpan · " + r.kodeKC + " versi " + r.versi));
      box.appendChild(el("div", "teks-kecil",
        "CKPN Net Flow " + rp(r.nfTotal) + " · Migration " + rp(r.migTotal) + " · PPKA " + rp(r.ppkaTotal) +
        " · " + r.jumlahPerubahan + " perubahan penyesuaian"));
      if (r.snapshot) {
        var b = el("button", "tautan", "Buka snapshot .xlsx");
        b.type = "button";
        b.addEventListener("click", function () { panggil("bukaSnapshot", { runId: r.runId }).catch(function (e) { alert(e.message); }); });
        box.appendChild(b);
      }
      if (r.pesanSnapshot) box.appendChild(el("div", "peringatan-box", r.pesanSnapshot));
      var ke = el("button", "tombol-sekunder lebar", "Lihat grup berikutnya (tab Per grup)");
      ke.type = "button";
      ke.addEventListener("click", function () {
        tampil("kartu-simpan", false);
        tampil("kartu-langkah", true);
        document.querySelector('.tab [data-tab="periode"]').click();
      });
      box.appendChild(ke);
      (r.pesan || []).forEach(function (p) {
        var m = el("div", "pesan-modul" + (p.peringatan ? " peringatan" : ""));
        m.appendChild(el("div", "judul-pesan", p.judul));
        m.appendChild(el("pre", "", p.pesan));
        box.appendChild(m);
      });
      w.appendChild(box);
      $("simpan-perubahan").innerHTML = "";
      $("simpan-catatan").value = "";
    }).catch(function (e) {
      $("simpan-hasil").innerHTML = "";
      $("simpan-hasil").appendChild(el("div", "peringatan-box", "Gagal menyimpan: " + e.message));
      btn.disabled = false;
    }).then(function () { btn.textContent = "Simpan"; });
  });

  $("btn-ke-simpan").addEventListener("click", bukaSimpan);
  $("btn-simpan-periksa").addEventListener("click", periksaSimpan);
  $("btn-simpan-batal").addEventListener("click", function () {
    tampil("kartu-simpan", false);
    tampil("kartu-langkah", true);
  });

  // =====================================================================
  // 6b. Hitung ulang Summary saja (setelah edit, sebelum simpan)
  // =====================================================================
  $("btn-refresh-summary").addEventListener("click", function () {
    var b = this, w = $("refresh-hasil");
    b.disabled = true;
    b.textContent = "Menghitung…";
    w.innerHTML = "";
    panggil("refreshSummary", {}, 600000).then(function (r) {
      var box = el("div", "angka-grup");
      [["CKPN Net Flow", r.nfTotal], ["CKPN Migration", r.migTotal], ["PPKA grup", r.ppkaGrup]].forEach(function (x) {
        var c = el("div", "", x[0]);
        c.appendChild(el("b", "", rp(x[1])));
        box.appendChild(c);
      });
      w.appendChild(box);
    }).catch(function (e) {
      w.appendChild(el("div", "peringatan-box", e.message));
    }).then(function () { b.disabled = false; b.textContent = "Hitung ulang Summary"; });
  });

  // =====================================================================
  // 7. Tab Periode: alur, grup, versi, konsolidasi, susunan
  // =====================================================================
  var statusP = null;   // hasil statusPeriode terakhir

  var LABEL_STATUS = {
    "belum": "Belum dihitung",
    "dihitung": "Dihitung — belum disimpan",
    "tersimpan": "Tersimpan",
    "dihitung-ulang": "Dihitung ulang — belum disimpan"
  };

  function muatPeriode(tanggal) {
    teks("periode-info", "Memuat…");
    panggil("statusPeriode", { tanggal: tanggal || "" }).then(function (d) {
      statusP = d;
      gambarPilihPeriode(d);
      gambarAlur(d);
      gambarGrup(d);
      gambarKonsolidasi(d);
    }).catch(function (e) { teks("periode-info", e.message); });
  }

  function gambarPilihPeriode(d) {
    var sel = $("pilih-periode");
    sel.innerHTML = "";
    var ada = {};
    (d.daftarPeriode || []).forEach(function (p) { ada[p.tanggal] = p; });
    if (d.tanggalMaster && !ada[d.tanggalMaster]) ada[d.tanggalMaster] = { tanggal: d.tanggalMaster, status: "Terbuka", jumlahGrup: 0 };
    Object.keys(ada).sort().reverse().forEach(function (t) {
      var p = ada[t];
      var o = el("option", "", t + (t === d.tanggalMaster ? " (Master)" : "") + " · " + p.jumlahGrup + " grup" +
                                    (p.status === "Final" ? " · Final" : ""));
      o.value = t;
      if (t === d.tanggal) o.selected = true;
      sel.appendChild(o);
    });
  }
  $("pilih-periode").addEventListener("change", function () { muatPeriode(this.value); });

  function gambarAlur(d) {
    var final = d.status === "Final";
    var tahap = [
      ["Susunan grup", d.susunanDitetapkan],
      ["Hitung & simpan (" + d.grupTersimpan + "/" + d.jumlahGrup + ")", d.siapKonsolidasi],
      ["Konsolidasi vs PPKA", final],
      ["Final", final]
    ];
    var ol = $("periode-alur");
    ol.innerHTML = "";
    var aktifSudah = false;
    tahap.forEach(function (t) {
      var cls = t[1] ? "selesai" : (!aktifSudah ? "aktif" : "");
      if (!t[1]) aktifSudah = true;
      var li = el("li", cls);
      li.appendChild(el("div", "bar-alur"));
      li.appendChild(el("span", "", t[0]));
      ol.appendChild(li);
    });

    var info = $("periode-info");
    info.innerHTML = "";
    var baris = [];
    if (final) baris.push("Periode dikunci oleh " + d.dikunciOleh + " · " + d.dikunciWaktu);
    if (!d.periodeMaster) baris.push("Hanya-baca: Master!C4 berisi periode " + (d.tanggalMaster || "—") + ". Ubah Master!C4 untuk menghitung periode ini.");
    else baris.push("Workbook saat ini: " + (d.kodeKCMaster || "tidak ada KC dicentang"));
    if (!d.susunanDitetapkan) baris.push("Susunan grup tahun " + d.tahun + " belum ditetapkan — sementara ditampilkan satu grup per kode.");
    else if (d.susunanTahun !== d.tahun) baris.push("Memakai susunan grup tahun " + d.susunanTahun + ".");
    if (d.susunanDitetapkan) baris.push(d.metode
      ? "Metode konsolidasi " + d.tahun + ": " + NAMA_METODE[d.metode] + (d.metodeTahun && d.metodeTahun !== d.tahun ? " (mengikuti " + d.metodeTahun + ")" : "")
      : "Metode konsolidasi belum ditetapkan — buka Atur susunan grup.");
    if (d.infoPengirim && d.bolehMenulis) baris.push(d.infoPengirim);
    if (!d.bolehMenulis) baris.push(d.infoPengirim);
    baris.forEach(function (b) { info.appendChild(el("div", "", b)); });
  }

  function gambarGrup(d) {
    var w = $("grup-isi");
    w.innerHTML = "";
    var final = d.status === "Final";
    var bolehHitung = d.periodeMaster && !final;

    (d.grup || []).forEach(function (g) {
      var k = el("div", "kartu-grup" + (g.diWorkbook ? " di-workbook" : ""));
      var head = el("div", "kepala-grup");
      var kiri = el("div");
      kiri.appendChild(el("div", "nama-grup", g.nama));
      kiri.appendChild(el("div", "teks-kecil", g.kc.join(", ") + (g.diWorkbook ? " · sedang di workbook" : "")));
      head.appendChild(kiri);
      var labelStatus = LABEL_STATUS[g.status] + (g.aktif ? " v" + g.aktif.versi : "");
      head.appendChild(el("span", "status-grup " + g.status, labelStatus));
      k.appendChild(head);

      if (g.aktif) {
        var rg = g.aktif.ringkasan || {};
        var angka = el("div", "angka-grup");
        [["Net Flow", rg.nf_total], ["Migration", rg.mig_total], ["PPKA", g.aktif.ppkaGrup]].forEach(function (x) {
          var c = el("div", "", x[0]); c.appendChild(el("b", "", rp(x[1]))); angka.appendChild(c);
        });
        k.appendChild(angka);
        k.appendChild(el("div", "teks-kecil", "Disimpan " + g.aktif.waktu + " oleh " + g.aktif.pengguna));
      } else if (g.dihitungDiPC) {
        k.appendChild(el("div", "teks-kecil", "Dihitung " + g.waktuHitung + " di PC ini"));
      }

      // ---- aksi ----
      var aksi = el("div", "aksi-grup");
      if (bolehHitung) {
        if (g.status === "dihitung" || g.status === "dihitung-ulang") {
          aksi.appendChild(tombol("Review & simpan", "tombol-utama", bukaSimpan));
          aksi.appendChild(tombol("Hitung ulang", "tombol-sekunder", function () { hitungGrup(g); }));
        } else if (g.status === "tersimpan" && g.diWorkbook) {
          aksi.appendChild(tombol("Edit & simpan ulang", "tombol-utama", bukaSimpan));
          aksi.appendChild(tombol("Hitung ulang", "tombol-sekunder", function () { hitungGrup(g); }));
        } else if (g.status === "tersimpan") {
          var b = tombol("Buka untuk diedit", "tombol-sekunder", function () { hitungGrup(g, true); });
          b.title = "Menghitung ulang grup ini dengan penyesuaian tersimpan, sehingga sheet kembali seperti saat disimpan";
          aksi.appendChild(b);
        } else {
          aksi.appendChild(tombol("Hitung grup ini", "tombol-utama", function () { hitungGrup(g); }));
        }
      }
      if (aksi.childNodes.length) k.appendChild(aksi);

      if (g.versi && g.versi.length) k.appendChild(daftarVersi(g.versi, d, final));
      w.appendChild(k);
    });

    if ((d.runLuarSusunan || []).length) {
      var luar = el("div", "kartu-grup");
      luar.appendChild(el("div", "nama-grup", "Di luar susunan grup"));
      luar.appendChild(el("div", "teks-kecil", "Kiriman dengan kombinasi KC yang tidak ada di susunan — tidak ikut konsolidasi."));
      luar.appendChild(daftarVersi(d.runLuarSusunan, d, final, true));
      w.appendChild(luar);
    }
  }

  function tombol(label, cls, fn) {
    var b = el("button", cls, label);
    b.type = "button";
    b.addEventListener("click", fn);
    return b;
  }

  function daftarVersi(list, d, final, terbuka) {
    var det = el("details", "versi-grup");
    if (terbuka) det.open = true;
    det.appendChild(el("summary", "", "Riwayat versi (" + list.length + ")"));
    list.forEach(function (v) {
      var row = el("div", "baris-versi" + (v.dihapus ? " terhapus" : ""));
      var rg = v.ringkasan || {};
      var ket = el("div");
      ket.appendChild(el("span", "", (v.kodeKC && terbuka ? v.kodeKC + " · " : "") + "v" + v.versi + (v.aktif ? " (aktif)" : "") +
        " · NF " + rp(rg.nf_total) + " · Mig " + rp(rg.mig_total)));
      ket.appendChild(el("div", "teks-kecil", v.pengguna + " · " + v.waktu +
        (v.dihapus ? " · dihapus " + v.penggunaHapus + ": " + v.alasanHapus : (v.catatan ? " · " + v.catatan : ""))));
      row.appendChild(ket);
      var aksi = el("div", "aksi-versi");
      if (v.adaSnapshot) aksi.appendChild(tombol("xlsx", "", function () {
        panggil("bukaSnapshot", { runId: v.id }).catch(function (e) { alert(e.message); });
      }));
      if (!v.dihapus && !final && d.bolehMenulis) {
        var h = tombol("Hapus", "hapus", function () {
          var alasan = window.prompt("Hapus versi v" + v.versi + " (" + v.kodeKC + ")?\nAlasan (wajib):");
          if (!alasan) return;
          panggil("hapusVersi", { runId: v.id, alasan: alasan })
            .then(function () { muatPeriode(d.tanggal); })
            .catch(function (e) { alert(e.message); });
        });
        aksi.appendChild(h);
      }
      row.appendChild(aksi);
      det.appendChild(row);
    });
    return det;
  }

  function hitungGrup(g, untukEdit) {
    var pesan = (untukEdit ? "Buka " + g.nama + " untuk diedit.\n\n" : "Hitung " + g.nama + ".\n\n") +
      "Centang KC di Master akan diubah menjadi: " + g.kc.join(", ") + "\n" +
      "Semua langkah (Individu s.d. Summary) dijalankan dengan penyesuaian tersimpan.\n" +
      "Isi sheet hasil saat ini akan ditimpa. Lanjutkan?";
    if (!window.confirm(pesan)) return;
    panggil("hitungGrup", { kodeKC: g.kodeKC, terapkanPenyesuaian: true }).then(function () {
      window.scrollTo(0, 0);   // progres tampil di area proses (atas)
    }).catch(function (e) { alert(e.message); });
  }

  // ---------------- Konsolidasi (Tahap 3c: satu metode + usulan jurnal) ----------------
  var NAMA_METODE = { nf: "Net Flow", mig: "Migration" };
  var LABEL_JENIS = {
    transisi: "Jurnal transisi — 1x, awal 2027",
    reguler:  "Jurnal reguler"
  };

  function persen(a, b) { return b ? (a / b * 100).toFixed(1).replace(".", ",") + "%" : "—"; }

  // "1.234.567" / "1234567" / "1.234.567,5" → angka; kosong/invalid → null
  function parseAngka(t) {
    t = String(t || "").replace(/\s/g, "").replace(/\./g, "").replace(",", ".");
    if (t === "" || t === "-") return null;
    var v = Number(t);
    return isFinite(v) ? v : null;
  }

  function gambarKonsolidasi(d) {
    var w = $("konsolidasi-isi");
    w.innerHTML = "";
    var final = d.status === "Final";
    if (!d.siapKonsolidasi && !final) {
      w.appendChild(el("div", "", d.susunanDitetapkan
        ? "Tersedia setelah semua grup tersimpan (" + d.grupTersimpan + "/" + d.jumlahGrup + ")."
        : "Tetapkan susunan grup terlebih dahulu."));
      return;
    }
    w.appendChild(el("div", "", "Memuat…"));
    panggil("konsolidasi", { tanggal: d.tanggal }).then(function (k) {
      w.innerHTML = "";
      gambarMetode(w, k);
      gambarTabelKonsolidasi(w, k);
      if (k.jurnal) gambarJurnal(w, k, d, final);
      gambarAksiKonsolidasi(w, k, d, final);
    }).catch(function (e) { w.textContent = e.message; });
  }

  function gambarMetode(w, k) {
    var b = el("div", "metode-baris");
    if (k.metodeCampuran) {
      w.appendChild(el("div", "peringatan-box",
        "Periode ini dikunci sebelum Tahap 3c dengan metode per grup. Angka ditampilkan sesuai keputusan saat itu."));
      return;
    }
    if (!k.metode) {
      var p = el("div", "peringatan-box", "Metode konsolidasi tahun " + k.tanggal.substring(0, 4) +
        " belum ditetapkan. Metode dipilih sekali setahun di susunan grup, lalu berlaku untuk semua grup.");
      var t = tombol("Atur susunan grup", "tautan", function () { $("btn-atur-susunan").click(); });
      p.appendChild(t);
      w.appendChild(p);
      return;
    }
    b.appendChild(el("span", "chip-metode", "Metode: " + NAMA_METODE[k.metode]));
    b.appendChild(el("span", "teks-kecil", "ditetapkan di susunan grup " + k.metodeTahun + " · berlaku untuk semua grup"));
    w.appendChild(b);
  }

  function gambarTabelKonsolidasi(w, k) {
    var t = el("table", "tabel-konsolidasi");
    var h = el("tr");
    ["Grup", "Net Flow", "Migration", "PPKA"].forEach(function (x) { h.appendChild(el("th", "", x)); });
    t.appendChild(h);

    k.grup.forEach(function (g) {
      var tr = el("tr");
      var c1 = el("td");
      c1.appendChild(el("div", "", g.nama));
      c1.appendChild(el("div", "teks-kecil", "v" + g.versi));
      tr.appendChild(c1);
      tr.appendChild(el("td", g.dipakai === "nf" ? "dipakai" : "", rp(g.nf)));
      tr.appendChild(el("td", g.dipakai === "mig" ? "dipakai" : "", rp(g.mig)));
      tr.appendChild(el("td", "", rp(g.ppka)));
      t.appendChild(tr);
    });

    var trAba = el("tr");
    trAba.appendChild(el("td", "", "ABA (KC0500)"));
    trAba.appendChild(el("td", "", rp(k.aba.ckpn)));
    trAba.appendChild(el("td", "", rp(k.aba.ckpn)));
    trAba.appendChild(el("td", "", rp(k.aba.ppka)));
    t.appendChild(trAba);

    var trTot = el("tr", "total");
    trTot.appendChild(el("td", "", "Total"));
    trTot.appendChild(el("td", k.metode === "nf" ? "dipakai" : "", rp(k.pembanding.nf)));
    trTot.appendChild(el("td", k.metode === "mig" ? "dipakai" : "", rp(k.pembanding.mig)));
    trTot.appendChild(el("td", "", rp(k.total.ppka)));
    t.appendChild(trTot);
    var gulir = el("div", "gulir-x");
    gulir.appendChild(t);
    w.appendChild(gulir);
    w.appendChild(el("div", "teks-kecil", "Kolom yang disorot = metode terpakai; total CKPN " + rp(k.total.ckpn) + "."));

    if (k.total.ckpn !== null && k.total.ckpn !== undefined) {
      var sel = k.total.ckpn - k.total.ppka;
      w.appendChild(el("div", "teks-kecil" + (sel < 0 ? " kurang" : ""),
        "CKPN " + (sel >= 0 ? "di atas" : "di bawah") + " PPKA sebesar " + rp(Math.abs(sel)) +
        " (" + persen(k.total.ckpn, k.total.ppka) + " dari PPKA)"));
    }
    if (k.metode && k.pembanding.terendah !== k.metode)
      w.appendChild(el("div", "teks-kecil", "Sebagai pembanding: periode ini " + NAMA_METODE[k.pembanding.terendah] +
        " lebih rendah " + rp(Math.abs(k.pembanding.nf - k.pembanding.mig)) + ". Metode tetap mengikuti susunan tahunan."));
    if (!k.aba.tersedia)
      w.appendChild(el("div", "teks-kecil", "CKPN ABA (Summary!C26) belum tersimpan di kiriman — simpan ulang salah satu grup untuk melengkapinya."));
  }

  function gambarJurnal(w, k, d, final) {
    var sek = el("div", "blok-jurnal");
    k.jurnal.forEach(function (j) {
      var kepala = el("div", "baris-antara");
      kepala.appendChild(el("h4", "", "Usulan jurnal CKPN"));
      var lencana = el("div", "lencana-jurnal");
      lencana.appendChild(el("span", "jenis-jurnal " + j.jenis, LABEL_JENIS[j.jenis] || j.jenis));
      if (j.simulasi) lencana.appendChild(el("span", "jenis-jurnal simulasi", "simulasi"));
      kepala.appendChild(lencana);
      sek.appendChild(kepala);

      var box = el("div", "jurnal-item");
      if (j.komponen !== "total") box.appendChild(el("div", "nama-grup", j.label));
      var angka = el("div", "angka-grup");
      [["CKPN (metode + ABA)", j.ckpn], ["PPKA", j.ppka],
       [j.selisihPpka < 0 ? "PPKA > CKPN" : "PPKA < CKPN", Math.abs(j.selisihPpka)]]
        .forEach(function (x) { var c = el("div", "", x[0]); c.appendChild(el("b", "", rp(x[1]))); angka.appendChild(c); });
      box.appendChild(angka);

      if (j.debit) {
        var tj = el("table", "tabel-jurnal");
        [["Db.", j.debit], ["Kr.", j.kredit]].forEach(function (x, i) {
          var tr = el("tr", i === 1 ? "kredit" : "");
          tr.appendChild(el("td", "dk", x[0]));
          tr.appendChild(el("td", "", x[1]));
          tr.appendChild(el("td", "angka", rp(j.nominal)));
          tj.appendChild(tr);
        });
        box.appendChild(tj);
      }
      box.appendChild(el("div", "teks-kecil" + (j.arah > 0 ? " kurang" : ""), j.keterangan));
      sek.appendChild(box);
    });

    var aturan = el("details", "aturan-jurnal");
    aturan.appendChild(el("summary", "", "Aturan jurnal"));
    var t = el("table", "tabel-jurnal");
    [["", "Transisi (1x, awal 2027)", "Reguler"],
     ["PPKA < CKPN", "Db. Laba ditahan – Kr. CKPN", "Db. Biaya CKPN – Kr. CKPN"],
     ["PPKA > CKPN", "Db. CKPN – Kr. Laba ditahan", "Db. Cadangan CKPN – Kr. Biaya CKPN / Pendapatan"]]
      .forEach(function (r, i) {
        var tr = el("tr");
        r.forEach(function (c) { tr.appendChild(el(i === 0 ? "th" : "td", "", c)); });
        t.appendChild(tr);
      });
    aturan.appendChild(t);
    aturan.appendChild(el("p", "teks-kecil", "CKPN = total CKPN grup sesuai metode tahunan ditambah CKPN ABA; PPKA = total PPKA termasuk ABA. " +
      "Nominal = selisih penuh bulan berjalan; setelah dibukukan ke CBS, PPKA = CKPN, dan bulan berikutnya selisih dicek lagi dari PPKA OJK yang baru. " +
      "Jurnal transisi dibentuk sekali dari posisi Desember 2026; periode sesudahnya memakai jurnal reguler. " +
      "Periode sebelum Desember 2026 ditandai simulasi."));
    sek.appendChild(aturan);
    w.appendChild(sek);
  }

  function gambarAksiKonsolidasi(w, k, d, final) {
    if (final) {
      w.appendChild(el("div", "teks-kecil jarak-atas", "Ditetapkan oleh " + d.dikunciOleh + " · " + d.dikunciWaktu +
        (d.catatanKunci ? " · " + d.catatanKunci : "")));
      if (d.bolehMenulis) {
        w.appendChild(tombol("Buka kunci periode", "tombol-sekunder lebar", function () {
          var alasan = window.prompt("Alasan membuka kunci periode " + d.tanggal + ":");
          if (!alasan) return;
          panggil("bukaKunci", { tanggal: d.tanggal, alasan: alasan })
            .then(function () { muatPeriode(d.tanggal); }).catch(function (e) { alert(e.message); });
        }));
      }
      return;
    }
    if (!d.bolehMenulis || !k.metode) return;

    var lab = el("label", "label-isian", "Catatan penetapan (opsional)");
    lab.setAttribute("for", "konsolidasi-catatan");
    w.appendChild(lab);
    var ta = el("textarea");
    ta.id = "konsolidasi-catatan"; ta.rows = 2;
    ta.placeholder = "Mis. nomor memo / persetujuan direksi";
    w.appendChild(ta);
    w.appendChild(tombol("Tetapkan & kunci periode", "tombol-utama jarak-atas", function () {
      var ringkas = (k.jurnal || []).map(function (j) {
        return j.debit ? "Db " + j.debit + " / Kr " + j.kredit + " " + rp(j.nominal) : "tanpa jurnal";
      }).join("\n");
      if (!window.confirm("Kunci periode " + d.tanggal + " dengan metode " + NAMA_METODE[k.metode] + "?\n\n" +
                          (ringkas ? "Usulan jurnal yang disimpan:\n" + ringkas + "\n\n" : "") +
                          "Setelah dikunci, grup tidak bisa disimpan ulang atau dihapus sampai kunci dibuka.")) return;
      panggil("tetapkanPeriode", { tanggal: d.tanggal, catatan: ta.value })
        .then(function () { muatPeriode(d.tanggal); })
        .catch(function (e) { alert(e.message); });
    }));
  }

  // ---------------- Susunan grup ----------------
  $("btn-atur-susunan").addEventListener("click", function () {
    if (!statusP) return;
    teks("susunan-judul", "Susunan grup tahun " + statusP.tahun);
    var w = $("susunan-isi");
    w.innerHTML = "";
    var nama = {};
    (statusP.grup || []).forEach(function (g) { g.kc.forEach(function (k) { nama[k] = g.nama; }); });
    ["KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100"].forEach(function (kc) {
      var lab = el("label", "kode", kc);
      lab.setAttribute("for", "sg-" + kc);
      var inp = el("input", "isian");
      inp.type = "text"; inp.id = "sg-" + kc; inp.value = nama[kc] || kc;
      w.appendChild(lab); w.appendChild(inp);
    });
    $("susunan-dasar").value = statusP.susunanDasar || "";
    document.querySelectorAll('input[name="susunan-metode"]').forEach(function (r) {
      r.checked = r.value === statusP.metode;
    });
    $("susunan-hasil").innerHTML = "";
    $("btn-susunan-simpan").disabled = !statusP.bolehMenulis;
    tampil("kartu-susunan", true);
  });
  $("btn-susunan-perkc").addEventListener("click", function () {
    ["KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100"].forEach(function (kc) { $("sg-" + kc).value = kc; });
  });
  $("btn-susunan-batal").addEventListener("click", function () { tampil("kartu-susunan", false); });
  $("btn-susunan-simpan").addEventListener("click", function () {
    var peta = {};
    ["KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100"].forEach(function (kc) { peta[kc] = $("sg-" + kc).value.trim(); });
    var m = document.querySelector('input[name="susunan-metode"]:checked');
    if (!m) {
      $("susunan-hasil").innerHTML = "";
      $("susunan-hasil").appendChild(el("div", "peringatan-box", "Pilih metode konsolidasi (Net Flow atau Migration)."));
      return;
    }
    if (statusP.metode && m.value !== statusP.metode &&
        !window.confirm("Metode konsolidasi diganti dari " + NAMA_METODE[statusP.metode] + " ke " + NAMA_METODE[m.value] +
                        ". Perubahan ini tercatat di log aktivitas. Lanjutkan?")) return;
    panggil("simpanSusunan", {
      tahun: statusP.tahun, peta: peta, dasar: $("susunan-dasar").value,
      metode: m.value, kebijakanSaldo: "ckpn"
    }).then(function () {
      tampil("kartu-susunan", false);
      muatPeriode(statusP.tanggal);
    }).catch(function (e) {
      $("susunan-hasil").innerHTML = "";
      $("susunan-hasil").appendChild(el("div", "peringatan-box", e.message));
    });
  });

  // =====================================================================
  // 7b. Data penyesuaian (Tahap 3d): lihat, cari, tambah, ubah, hapus, riwayat
  //     Berlaku pada perhitungan berikutnya (tombol VBA maupun panel).
  // =====================================================================
  var dataPny = null;          // hasil daftarPenyesuaian
  var modulPny = "individu";   // tab aktif: individu | lgdcs
  var BATAS_TAMPIL = 300;

  var LABEL_JENIS_CS = { ubah: "Ubah nilai", hapus: "Dikecualikan", tambah: "Baris manual" };

  function muatPenyesuaian() {
    teks("penyesuaian-isi", "Memuat…");
    return panggil("daftarPenyesuaian").then(function (d) {
      dataPny = d;
      $("btn-pny-tambah").disabled = !d.bolehMenulis;
      gambarPenyesuaian();
    }).catch(function (e) { teks("penyesuaian-isi", e.message); });
  }

  function uraianPny(p) {
    if (modulPny === "individu")
      return "Agunan " + rp(p.jaminan) + " (sistem " + rp(p.jaminanSistem) + ") · biaya " + rp(p.biaya);
    if (p.jenis === "hapus") return "Dikecualikan dari LGD CS";
    if (p.jenis === "tambah") return "Baris manual · pokok " + rp(p.pokok) + " · agunan " + rp(p.nilaiAgunan) +
                                     " · realisasi " + (p.recovery === null ? "rumus" : rp(p.recovery));
    return "Diubah" + (p.nilaiAgunan !== null ? " · agunan " + rp(p.nilaiAgunan) : "") +
                      (p.recovery !== null ? " · realisasi " + rp(p.recovery) : "");
  }

  function gambarPenyesuaian() {
    var d = dataPny;
    if (!d) return;
    document.querySelectorAll("[data-pny]").forEach(function (b) {
      var aktif = b.getAttribute("data-pny") === modulPny;
      b.setAttribute("aria-pressed", aktif ? "true" : "false");
      var n = (b.getAttribute("data-pny") === "individu" ? d.individu : d.lgdcs).length;
      b.textContent = (b.getAttribute("data-pny") === "individu" ? "CKPN Individu" : "LGD CS") + " (" + n + ")";
    });

    var cari = $("pny-cari").value.trim().toLowerCase();
    var semua = modulPny === "individu" ? d.individu : d.lgdcs;
    var list = semua.filter(function (p) {
      if (!cari) return true;
      return (p.kunci + " " + (p.nama || "") + " " + (p.jenis || "") + " " + (p.alasan || "")).toLowerCase().indexOf(cari) >= 0;
    });
    list.sort(function (a, b) { return (b.waktu || "").localeCompare(a.waktu || ""); });

    var w = $("penyesuaian-isi");
    w.innerHTML = "";
    if (!semua.length) { w.textContent = "Belum ada penyesuaian " + (modulPny === "individu" ? "Individu" : "LGD CS") + " tersimpan."; return; }
    if (!list.length) { w.textContent = "Tidak ada yang cocok dengan pencarian."; return; }

    list.slice(0, BATAS_TAMPIL).forEach(function (p) {
      var row = el("button", "item-pny");
      row.type = "button";
      var kepala = el("div", "baris-antara");
      kepala.appendChild(el("span", "kunci", p.kunci));
      if (modulPny === "lgdcs") kepala.appendChild(el("span", "jenis-cs " + p.jenis, LABEL_JENIS_CS[p.jenis] || p.jenis));
      row.appendChild(kepala);
      if (p.nama) row.appendChild(el("div", "", p.nama));
      row.appendChild(el("div", "teks-kecil", uraianPny(p)));
      row.appendChild(el("div", "teks-kecil", p.pengguna + " · " + p.waktu +
        (p.dipakai ? " · dipakai " + p.dipakai.tanggal + " (" + p.dipakai.kodeKC + ")" : " · belum dipakai di kiriman aktif") +
        (p.alasan ? " · " + p.alasan : "")));
      row.addEventListener("click", function () { bukaEditorPny(p); });
      w.appendChild(row);
    });
    if (list.length > BATAS_TAMPIL)
      w.appendChild(el("p", "teks-kecil", "Menampilkan " + BATAS_TAMPIL + " dari " + list.length + " — persempit pencarian."));
  }

  document.querySelectorAll("[data-pny]").forEach(function (b) {
    b.addEventListener("click", function () {
      modulPny = b.getAttribute("data-pny");
      tampil("kartu-pny-editor", false);
      gambarPenyesuaian();
    });
  });
  $("pny-cari").addEventListener("input", gambarPenyesuaian);
  $("btn-muat-penyesuaian").addEventListener("click", muatPenyesuaian);
  $("btn-pny-tambah").addEventListener("click", function () { bukaEditorPny(null); });

  // ---------------- Editor ----------------
  var editPny = null;   // item yang sedang diedit (null = tambah baru)

  function isiAngka(id, v) { $(id).value = v === null || v === undefined ? "" : fmt.format(v); }
  function bacaAngka(id) { return parseAngka($(id).value); }

  function aturFieldCs() {
    var j = $("pe-jenis").value;
    tampil("pe-grup-pokok", j === "tambah");
    tampil("pe-grup-tahun", j === "tambah");
    tampil("pe-grup-nilai", j !== "hapus");
    teks("pe-ket-jenis", j === "hapus" ? "Rekening ini tidak ikut LGD CS pada perhitungan berikutnya." :
      j === "tambah" ? "Baris ditambahkan di bawah data LGD CS (font ungu) bila rekening tidak ada di hasil sistem." :
      "Nilai menggantikan hasil sistem. Kosongkan realisasi untuk tetap memakai rumus prefill.");
  }
  $("pe-jenis").addEventListener("change", aturFieldCs);

  function bukaEditorPny(p) {
    editPny = p;
    var ind = modulPny === "individu";
    teks("pe-judul", (p ? "Ubah" : "Tambah") + " penyesuaian " + (ind ? "CKPN Individu" : "LGD CS"));
    teks("pe-label-kunci", ind ? "No. kontrak" : "No. rekening");
    $("pe-kunci").value = p ? p.kunci : "";
    $("pe-kunci").readOnly = !!p;
    $("pe-nama").value = p ? (p.nama || "") : "";
    $("pe-nama").readOnly = ind;   // nama Individu berasal dari hasil hitung
    tampil("pe-grup-nama", !ind || !!(p && p.nama));
    tampil("pe-grup-ind", ind);
    tampil("pe-grup-cs", !ind);
    if (ind) {
      isiAngka("pe-jaminan", p ? p.jaminan : null);
      isiAngka("pe-biaya", p ? p.biaya : null);
      teks("pe-sistem", p ? "Nilai agunan sistem saat disimpan: " + rp(p.jaminanSistem) : "Nilai agunan sistem diambil dari perhitungan terakhir di PC ini (bila ada).");
    } else {
      $("pe-jenis").value = p ? p.jenis : "ubah";
      isiAngka("pe-agunan", p ? p.nilaiAgunan : null);
      isiAngka("pe-recovery", p ? p.recovery : null);
      isiAngka("pe-pokok", p ? p.pokok : null);
      $("pe-thn-serah").value = p ? (p.thnSerah || "") : "";
      $("pe-thn-eks").value = p ? (p.thnEks || "") : "";
      aturFieldCs();
    }
    $("pe-alasan").value = "";
    $("pe-hasil").innerHTML = "";
    $("pe-riwayat").innerHTML = "";
    var boleh = dataPny && dataPny.bolehMenulis;
    $("btn-pe-simpan").disabled = !boleh;
    tampil("btn-pe-hapus", !!p && boleh);
    tampil("btn-pe-riwayat", !!p);
    tampil("kartu-pny-editor", true);
    $("kartu-pny-editor").scrollIntoView({ block: "nearest" });
    (p ? (ind ? $("pe-jaminan") : $("pe-jenis")) : $("pe-kunci")).focus();
  }

  function pesanEditor(teksPesan, ok) {
    var w = $("pe-hasil");
    w.innerHTML = "";
    w.appendChild(el("div", ok ? "info-box" : "peringatan-box", teksPesan));
  }

  $("btn-pe-batal").addEventListener("click", function () { tampil("kartu-pny-editor", false); });

  $("btn-pe-simpan").addEventListener("click", function () {
    var ind = modulPny === "individu";
    var data;
    try {
      data = ind
        ? { kunci: $("pe-kunci").value.trim(), jaminan: bacaAngka("pe-jaminan"), biaya: bacaAngka("pe-biaya") }
        : { kunci: $("pe-kunci").value.trim(), jenis: $("pe-jenis").value, nama: $("pe-nama").value.trim(),
            nilaiAgunan: bacaAngka("pe-agunan"), recovery: bacaAngka("pe-recovery"), pokok: bacaAngka("pe-pokok"),
            thnSerah: $("pe-thn-serah").value.trim(), thnEks: $("pe-thn-eks").value.trim() };
    } catch (e) { pesanEditor(e.message); return; }

    var b = this;
    b.disabled = true;
    panggil("simpanPenyesuaian", { modul: modulPny, data: data, alasan: $("pe-alasan").value }).then(function (r) {
      var dipakai = editPny && editPny.dipakai;
      pesanEditor("Tersimpan: " + r.sesudah + ". Berlaku pada perhitungan berikutnya" +
        (dipakai ? " — hitung ulang grup " + dipakai.kodeKC + " agar hasilnya berubah." : "."), true);
      return muatPenyesuaian();
    }).catch(function (e) { pesanEditor(e.message); })
      .then(function () { b.disabled = false; });
  });

  $("btn-pe-hapus").addEventListener("click", function () {
    if (!editPny) return;
    var alasan = $("pe-alasan").value;
    if (!window.confirm("Hapus penyesuaian " + editPny.kunci + "?\nPerhitungan berikutnya kembali memakai nilai sistem" +
                        (modulPny === "lgdcs" && editPny.jenis === "hapus" ? " (rekening masuk lagi ke LGD CS)" : "") + ".")) return;
    panggil("hapusPenyesuaian", { modul: modulPny, kunci: editPny.kunci, alasan: alasan }).then(function () {
      tampil("kartu-pny-editor", false);
      return muatPenyesuaian();
    }).catch(function (e) { pesanEditor(e.message); });
  });

  var LABEL_AKSI_LOG = {
    "baru": "penyesuaian baru (simpan grup)", "ubah": "diubah (simpan grup)", "kembali": "kembali ke sistem",
    "sistem": "nilai sistem diperbarui", "hapus-baris": "dikecualikan (simpan grup)", "hapus-manual": "baris manual dihapus",
    "tambah-dari-panel": "ditambah dari panel", "ubah-dari-panel": "diubah dari panel", "hapus-dari-panel": "dihapus dari panel"
  };

  $("btn-pe-riwayat").addEventListener("click", function () {
    if (!editPny) return;
    var w = $("pe-riwayat");
    w.textContent = "Memuat…";
    panggil("riwayatPenyesuaian", { modul: modulPny, kunci: editPny.kunci }).then(function (list) {
      w.innerHTML = "";
      if (!list.length) { w.textContent = "Belum ada catatan."; return; }
      list.forEach(function (x) {
        var r = el("div", "baris-log");
        r.appendChild(el("div", "", (LABEL_AKSI_LOG[x.aksi] || x.aksi) + " · " + x.pengguna + " · " + x.waktu +
          (x.periode ? " · periode " + x.periode : "")));
        r.appendChild(el("div", "teks-kecil", (x.sebelum ? x.sebelum + " → " : "") + (x.sesudah || "") + (x.alasan ? " · " + x.alasan : "")));
        w.appendChild(r);
      });
    }).catch(function (e) { w.textContent = e.message; });
  });

  // =====================================================================
  // 7c. Kelola database (Tahap 3d): hapus data periode / kosongkan (simulasi)
  // =====================================================================
  var infoDb = null;

  function muatInfoDatabase() {
    var w = $("db-isi");
    w.textContent = "Memuat…";
    return panggil("infoDatabase").then(function (d) {
      infoDb = d;
      gambarInfoDatabase();
    }).catch(function (e) { w.textContent = e.message; });
  }

  function gambarInfoDatabase() {
    var d = infoDb, w = $("db-isi");
    w.innerHTML = "";
    if (!d.adaDatabase) { w.textContent = "Database belum ada."; return; }
    w.appendChild(el("div", "teks-kecil", (d.periode || []).length + " periode · " + d.penyesuaianIndividu +
      " penyesuaian Individu · " + d.penyesuaianLgdCs + " penyesuaian LGD CS · " + d.ukuranKB + " KB"));
    if (!d.bolehMenulis) { w.appendChild(el("div", "peringatan-box", d.infoPengirim)); }

    var t = el("table", "tabel-staging");
    var h = el("tr");
    ["Periode", "Status", "Versi", ""].forEach(function (x) { h.appendChild(el("th", "", x)); });
    t.appendChild(h);
    (d.periode || []).forEach(function (p) {
      var tr = el("tr");
      var c1 = el("td");
      c1.appendChild(el("div", "", p.tanggal));
      if (p.metode) c1.appendChild(el("div", "teks-kecil", NAMA_METODE[p.metode] || p.metode));
      tr.appendChild(c1);
      tr.appendChild(el("td", "", p.status + (p.adaJurnal ? " · jurnal" : "")));
      tr.appendChild(el("td", "angka", p.grupAktif + " aktif / " + p.jumlahVersi));
      var c4 = el("td");
      if (d.bolehMenulis)
        c4.appendChild(tombol("Hapus", "tautan hapus", function () { formHapusPeriode(p); }));
      tr.appendChild(c4);
      t.appendChild(tr);
    });
    if ((d.periode || []).length) w.appendChild(t);
    else w.appendChild(el("div", "teks-kecil", "Tidak ada data periode."));

    if (d.bolehMenulis)
      w.appendChild(tombol("Kosongkan database…", "tombol-sekunder lebar bahaya", formKosongkan));
  }

  // Form konfirmasi yang dipakai bersama oleh hapus periode & kosongkan
  function formBahaya(judul, uraian, opsi, kataKunci, aksi) {
    var f = $("db-form");
    f.innerHTML = "";
    f.appendChild(el("div", "judul-kartu", judul));
    uraian.forEach(function (u) { f.appendChild(el("p", "teks-kecil", u)); });
    var cek = {};
    opsi.forEach(function (o) {
      var lab = el("label", "opsi");
      var c = el("input");
      c.type = "checkbox"; c.checked = !!o.bawaan;
      lab.appendChild(c);
      lab.appendChild(el("span", "", o.label));
      f.appendChild(lab);
      cek[o.id] = c;
    });
    var l1 = el("label", "label-isian", "Alasan (opsional)");
    l1.setAttribute("for", "db-alasan");
    f.appendChild(l1);
    var alasan = el("input", "isian"); alasan.id = "db-alasan"; alasan.type = "text";
    alasan.placeholder = "Mis. simulasi metode Migration";
    f.appendChild(alasan);
    var l2 = el("label", "label-isian", "Ketik " + kataKunci + " untuk konfirmasi");
    l2.setAttribute("for", "db-konfirmasi");
    f.appendChild(l2);
    var konf = el("input", "isian"); konf.id = "db-konfirmasi"; konf.type = "text"; konf.autocomplete = "off";
    f.appendChild(konf);
    var baris = el("div", "baris-tombol jarak-atas");
    var bHapus = tombol("Hapus permanen", "tombol-utama bahaya", function () {
      bHapus.disabled = true;
      var nilai = {};
      Object.keys(cek).forEach(function (k) { nilai[k] = cek[k].checked; });
      aksi(konf.value.trim(), alasan.value, nilai).then(function (r) {
        f.innerHTML = "";
        f.appendChild(el("div", "info-box", "Selesai. Arsip sebelum penghapusan: " + r.arsip));
        muatInfoDatabase();
        muatPenyesuaian();
        if (statusP) muatPeriode();
      }).catch(function (e) {
        bHapus.disabled = false;
        f.appendChild(el("div", "peringatan-box", e.message));
      });
    });
    bHapus.disabled = true;
    konf.addEventListener("input", function () { bHapus.disabled = konf.value.trim() !== kataKunci; });
    baris.appendChild(bHapus);
    baris.appendChild(tombol("Batal", "tombol-sekunder", function () { f.innerHTML = ""; tampil("db-form", false); }));
    f.appendChild(baris);
    tampil("db-form", true);
    f.scrollIntoView({ block: "nearest" });
    konf.focus();
  }

  function formHapusPeriode(p) {
    formBahaya("Hapus data periode " + p.tanggal,
      ["Menghapus permanen semua versi grup, hasil Individu/LGD CS, konsolidasi, dan jurnal periode ini (" + p.jumlahVersi + " versi)." +
       (p.status === "Final" ? " Periode ini berstatus Final." : ""),
       "Salinan arsip database dibuat lebih dulu di folder library\\backup. Log aktivitas tetap disimpan."],
      [{ id: "penyesuaian", label: "Hapus juga penyesuaian yang tercatat pada periode ini (" + p.penyesuaianPeriode + ")", bawaan: false },
       { id: "snapshot", label: "Hapus juga file snapshot .xlsx periode ini", bawaan: true }],
      p.tanggal,
      function (konf, alasan, o) {
        return panggil("hapusDataPeriode", { tanggal: p.tanggal, konfirmasi: konf, alasan: alasan,
                                             hapusPenyesuaian: o.penyesuaian, hapusSnapshot: o.snapshot });
      });
  }

  function formKosongkan() {
    formBahaya("Kosongkan database",
      ["Menghapus permanen SEMUA periode, versi grup, hasil, konsolidasi, dan jurnal. Cocok untuk mengakhiri masa simulasi.",
       "Salinan arsip database dibuat lebih dulu di folder library\\backup. Log aktivitas tetap disimpan."],
      [{ id: "penyesuaian", label: "Hapus juga semua penyesuaian Individu & LGD CS", bawaan: false },
       { id: "susunan", label: "Hapus juga susunan grup & metode tahunan", bawaan: false },
       { id: "snapshot", label: "Hapus juga file snapshot .xlsx", bawaan: true }],
      infoDb.konfirmasiSemua,
      function (konf, alasan, o) {
        return panggil("kosongkanDatabase", { konfirmasi: konf, alasan: alasan, penyesuaian: o.penyesuaian,
                                              susunan: o.susunan, hapusSnapshot: o.snapshot });
      });
  }

  $("db-kelola").addEventListener("toggle", function () { if (this.open) muatInfoDatabase(); });


  // =====================================================================
  // 8. Tab Overview Data (Tahap 3c): OS, EAD, selisih, NPF per KC
  //    Sumber: file template di Master!D14 (dibaca oleh DataOverviewBuilder)
  // =====================================================================
  var dataOverview = null;
  var LABEL_KUALITAS = ["Lancar", "DPK", "Kurang Lancar", "Diragukan", "Macet"];
  var fmtJuta = new Intl.NumberFormat("id-ID", { maximumFractionDigits: 0 });

  function nilaiOv(v) {
    if (v === null || v === undefined) return "—";
    return $("ov-satuan").value === "juta" ? fmtJuta.format(v / 1e6) : fmt.format(v);
  }
  function persenOv(v) { return v === null || v === undefined ? "—" : (v * 100).toFixed(2).replace(".", ",") + "%"; }

  function muatOverview() {
    var b = $("btn-ov-muat");
    b.disabled = true;
    b.textContent = "Membaca template…";
    teks("ov-info", "Membuka file template (Master!D14) — beberapa detik.");
    panggil("overviewData", {}, 600000).then(function (d) {
      dataOverview = d;
      gambarOverview();
    }).catch(function (e) {
      $("ov-info").textContent = e.message;
    }).then(function () { b.disabled = false; b.textContent = dataOverview ? "Muat ulang" : "Muat data"; });
  }

  function namaGrupKC(kode) {
    var nama = "";
    ((statusP && statusP.grup) || []).forEach(function (g) {
      if (g.kc.indexOf(kode) >= 0 && g.nama !== kode) nama = g.nama;
    });
    return nama;
  }

  function gambarOverview() {
    var d = dataOverview;
    if (!d) return;
    teks("ov-info", (d.namaBPR ? d.namaBPR + " · " : "") + "posisi " + (d.periode || "—") + " · " + d.fileTemplate +
                    " · dimuat " + d.waktu);
    teks("ov-satuan-ket", $("ov-satuan").value === "juta" ? "Angka dalam juta Rp" : "Angka dalam Rp");

    // ---- ringkasan keuangan ----
    var k = $("ov-keuangan");
    k.innerHTML = "";
    var t = d.total;
    var kartu = [
      ["Total OS pembiayaan", t.totalOS], ["Total EAD", t.totalEAD],
      ["Selisih EAD − OS", t.selisih], ["NPF (gross)", null, persenOv(t.npf)]
    ];
    if (d.keuangan) kartu = kartu.concat([
      ["Aset", d.keuangan.asset], ["OS neraca (GB0200)", d.keuangan.osNeraca],
      ["CKPN neraca", d.keuangan.ckpnNeraca], ["PPKA template", t.totalPPKA],
      ["Laba tahun lalu", d.keuangan.labaLalu], ["Laba berjalan", d.keuangan.labaBerjalan]
    ]);
    kartu.forEach(function (x) {
      var c = el("div", "", x[0]);
      c.appendChild(el("b", "", x[2] !== undefined ? x[2] : nilaiOv(x[1])));
      k.appendChild(c);
    });
    var rekon = $("ov-rekon");
    rekon.textContent = "";
    if (d.keuangan) {
      var beda = d.keuangan.osNeraca - t.totalOS;
      rekon.textContent = Math.abs(beda) < 1 ? "OS segmen sama dengan OS neraca GB0200."
        : "OS neraca GB0200 " + (beda > 0 ? "lebih besar " : "lebih kecil ") + nilaiOv(Math.abs(beda)) +
          " dari total OS segmen (cek pos yang tidak masuk KC0600–KC1100).";
    }

    // ---- tabel per KC ----
    var w = $("ov-tabel");
    w.innerHTML = "";
    var tb = el("table", "tabel-konsolidasi tabel-ov");
    var h = el("tr");
    ["KC", "OS", "EAD", "EAD − OS", "NPF"].forEach(function (x) { h.appendChild(el("th", "", x)); });
    tb.appendChild(h);
    d.segmen.concat([d.total]).forEach(function (s) {
      var total = s.kode === "Total";
      var tr = el("tr", total ? "total" : "baris-klik");
      var c1 = el("td");
      c1.appendChild(el("div", "", s.kode));
      var sub = total ? s.debitur + " debitur" : (s.ada ? (namaGrupKC(s.kode) || s.debitur + " debitur") : s.pesan);
      c1.appendChild(el("div", "teks-kecil", sub));
      tr.appendChild(c1);
      tr.appendChild(el("td", "", nilaiOv(s.totalOS)));
      tr.appendChild(el("td", "", nilaiOv(s.totalEAD)));
      tr.appendChild(el("td", s.selisih < 0 ? "kurang" : "", nilaiOv(s.selisih)));
      tr.appendChild(el("td", "", persenOv(s.npf)));
      tb.appendChild(tr);

      if (!s.ada) return;
      var trD = el("tr", "rinci");
      trD.hidden = true;
      var td = el("td");
      td.colSpan = 5;
      td.appendChild(tabelKualitas(s));
      trD.appendChild(td);
      tb.appendChild(trD);
      tr.tabIndex = 0;
      tr.setAttribute("aria-expanded", "false");
      function buka() {
        trD.hidden = !trD.hidden;
        tr.setAttribute("aria-expanded", trD.hidden ? "false" : "true");
      }
      tr.addEventListener("click", buka);
      tr.addEventListener("keydown", function (e) { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); buka(); } });
    });
    w.appendChild(tb);
    w.appendChild(el("p", "teks-kecil catatan", "Klik baris untuk rincian per kualitas. NPF = (Kurang Lancar + Diragukan + Macet) ÷ OS. " +
      "KC0600–KC0900: EAD = OS. KC1000: OS saldo modal (AG), EAD tunggakan pokok + basil (AL+AN). " +
      "KC1100: OS nilai kontrak − penyusutan (AB−AI), EAD tunggakan pokok + ujroh (AL+AM)."));

    var log = $("ov-log");
    log.innerHTML = "";
    (d.log || []).forEach(function (x) { log.appendChild(el("div", "peringatan-box", x)); });
  }

  function tabelKualitas(s) {
    var t = el("table", "tabel-kualitas");
    var h = el("tr");
    ["Kualitas", "OS", "EAD", "PPKA"].forEach(function (x) { h.appendChild(el("th", "", x)); });
    t.appendChild(h);
    LABEL_KUALITAS.forEach(function (lbl, i) {
      var tr = el("tr");
      tr.appendChild(el("td", "", (i + 1) + " · " + lbl));
      tr.appendChild(el("td", "", nilaiOv(s.os[i])));
      tr.appendChild(el("td", "", nilaiOv(s.ead[i])));
      tr.appendChild(el("td", "", nilaiOv(s.ppka[i])));
      t.appendChild(tr);
    });
    var trT = el("tr", "total");
    trT.appendChild(el("td", "", "NPF " + persenOv(s.npf)));
    trT.appendChild(el("td", "", nilaiOv(s.totalOS)));
    trT.appendChild(el("td", "", nilaiOv(s.totalEAD)));
    trT.appendChild(el("td", "", nilaiOv(s.totalPPKA)));
    t.appendChild(trT);
    return t;
  }

  $("btn-ov-muat").addEventListener("click", muatOverview);
  $("ov-satuan").addEventListener("change", gambarOverview);


  // =====================================================================
  // Diagnostik
  // =====================================================================
  function muatDiagnostik() {
    return panggil("ping").then(function (d) {
      $("d-jembatan").innerHTML = '<span class="ok">Terhubung</span>';
      teks("lencana-versi", "v" + d.versiAddin);
      teks("d-excel", d.versiExcel + " · " + d.arsitektur);
      teks("d-webview", d.webview2);
      teks("d-user", d.user + " @ " + d.komputer);
      teks("d-library", d.folderLibrary);
      teks("d-log", d.fileLog);
      teks("d-db", (d.databaseAda ? "Ada · " : "Belum ada (dibuat saat Simpan grup pertama) · ") + d.fileDatabase);
    }).catch(function (err) {
      $("d-jembatan").innerHTML = '<span class="err">Gagal</span> ' + err.message;
      $("diag").open = true;
    });
  }

  document.querySelectorAll("[data-folder]").forEach(function (b) {
    b.addEventListener("click", function () {
      panggil("bukaFolder", { folder: b.getAttribute("data-folder") }).catch(function () {});
    });
  });

  muatDiagnostik();
  muatPersiapan();
})();
