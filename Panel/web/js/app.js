// =====================================================================
// Panel CKPN — logika sisi panel
// Tab (Tahap 5f): Ringkasan · Hitung (Per grup / Manual / Penyesuaian) · Tetapkan · Analisis (Tren / Komposisi / Versi) · ⚙ Pengaturan
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

  // ---------------- Tab (Tahap 5f) ----------------
  //   Ringkasan · Hitung (Per grup / Manual / Penyesuaian) · Tetapkan · Analisis (Tren / Komposisi / Versi) · ⚙ Pengaturan
  //   Periode yang ditampilkan dipilih sekali di header (#pilih-periode) dan berlaku untuk semua tab.
  var tombolTab = document.querySelectorAll('.tab [role="tab"]');
  var tabAktif = "ringkasan", tabSebelum = "ringkasan";
  var subAktif = { hitung: "pergrup", analisis: "tren" };

  // Periode terpilih untuk dimuat ulang: "" = ikuti Master!C4 (bila yang tampil memang periode Master)
  function periodeTerpilih() {
    if (!statusP) return "";
    return statusP.tanggal === statusP.tanggalMaster ? "" : statusP.tanggal;
  }

  function muatSub(grup, nama) {
    if (grup === "hitung") {
      if (nama === "penyesuaian") muatPenyesuaian();
      if (nama === "manual" && !dataLangkah.length) muatPersiapan();
    } else if (grup === "analisis") {
      if (nama === "tren") muatRiwayat();
      if (nama === "komposisi") {
        muatSumberAnalisis();
        muatPenjelasan($("pj-periode").value || (statusP ? statusP.tanggal : ""));
      }
      if (nama === "versi" && $("kartu-log").open) { muatLog($("lp-bulan").value); muatInfoAuditLog(); }
    }
  }

  function pindahSub(grup, nama, tanpaMuat) {
    subAktif[grup] = nama;
    document.querySelectorAll('.sub-nav[data-grup="' + grup + '"] [data-sub]').forEach(function (b) {
      var aktif = b.getAttribute("data-sub") === nama;
      b.setAttribute("aria-pressed", aktif ? "true" : "false");
      $("sub-" + b.getAttribute("data-sub")).hidden = !aktif;
    });
    Grafik.sembunyiTip();
    if (!tanpaMuat) muatSub(grup, nama);
  }

  function pindahTab(nama, sub, tanpaMuat) {
    if (nama !== tabAktif && tabAktif !== "pengaturan") tabSebelum = tabAktif;
    tabAktif = nama;
    tombolTab.forEach(function (x) {
      var aktif = x.getAttribute("data-tab") === nama;
      x.setAttribute("aria-selected", aktif ? "true" : "false");
      $("tab-" + x.getAttribute("data-tab")).hidden = !aktif;
    });
    Grafik.sembunyiTip();
    if (sub) pindahSub(nama, sub, true);
    if (tanpaMuat) return;
    if (nama === "ringkasan") { muatPeriode(periodeTerpilih()); if (!dataOverview) muatOverview(); }
    if (nama === "hitung" || nama === "tetapkan") muatPeriode(periodeTerpilih());
    if (nama === "pengaturan") {
      muatPeriode(periodeTerpilih());
      muatDiagnostik();
      if ($("db-kelola").open) muatInfoDatabase();
    }
    if (nama === "analisis" && subAktif.analisis === "versi") muatPeriode(periodeTerpilih());
    if (subAktif[nama]) muatSub(nama, subAktif[nama]);
  }

  tombolTab.forEach(function (b) {
    b.addEventListener("click", function () { pindahTab(b.getAttribute("data-tab")); });
  });
  document.querySelectorAll(".sub-nav [data-sub]").forEach(function (b) {
    b.addEventListener("click", function () {
      pindahSub(b.parentNode.getAttribute("data-grup"), b.getAttribute("data-sub"));
    });
  });
  $("btn-pengaturan-kembali").addEventListener("click", function () { pindahTab(tabSebelum || "ringkasan"); });

  // =====================================================================
  // 1. Persiapan: baca Master & validasi setiap langkah
  // =====================================================================
  var dataLangkah = [];     // hasil siapkanRun
  var infoTopN = null;      // Top-N Master!C10 vs ketetapan tahunan (dari siapkanRun)

  function muatPersiapan() {
    teks("periode", "Memeriksa…");
    $("btn-jalankan").disabled = true;
    return panggil("siapkanRun").then(function (d) {
      teks("periode", d.periode || "Master!C4 belum diisi");
      infoTopN = d;
      var ti = $("topn-info");
      ti.innerHTML = "";
      if (d.topNTahun && d.topNMaster !== d.topNTahun)
        ti.appendChild(el("div", "peringatan-box", "Top-N di Master!C10 = " + d.topNMaster + ", sedangkan ketetapan grup " + d.namaGrup + " tahun " + d.tahun +
          " = " + d.topNTahun + ". Hitung manual tetap bisa dijalankan, tetapi hasilnya tidak dapat disimpan ke staging. " +
          "Hitung dari Hitung › Per grup untuk memakai Top-N tahunan secara otomatis."));
      else if (d.topNTahun)
        ti.appendChild(el("div", "teks-kecil", "Top-N " + d.topNMaster + " sesuai ketetapan grup " + d.namaGrup + " tahun " + d.tahun + "."));
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

  // Status kartu grup (Hitung › Per grup) disegarkan setiap kali isi workbook/database berubah:
  // selesai hitung, selesai simpan, dan saat kartu hasil/simpan ditutup — tanpa perlu pindah tab.
  function segarkanPerGrup() {
    if (!statusP) return;   // periode belum pernah dimuat
    muatPeriode(statusP.tanggal === statusP.tanggalMaster ? "" : statusP.tanggal);
  }

  onEvent("runSelesai", function (d) {
    if (batchAktif) {
      // Saat hitung semua grup, hasil per grup dirangkum di kartu batch; kartu hasil hanya
      // ditampilkan untuk grup tempat proses berhenti (lihat batchSelesai).
      runTerakhirBatch = d;
      tampil("kartu-progres", false);
      run = null;
      segarkanPerGrup();
      return;
    }
    tampilkanHasil(d);
    if (d.status === "selesai") muatReview();
    run = null;
    muatPersiapan();   // segarkan status langkah untuk run berikutnya
    segarkanPerGrup();
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
    segarkanPerGrup();
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

    if (cs && cs.lewati) {
      var b0 = el("div", "blok-review");
      b0.appendChild(el("h4", "", "LGD Collateral Shortfall"));
      b0.appendChild(el("div", "teks-kecil", "Tidak dihitung bulan ini — " + (cs.keterangan || "LGD memakai acuan Desember") + "."));
      wadah.appendChild(b0);
      cs = null;
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
      if (d.acuan) pw.appendChild(el("div", "info-box", d.acuan));
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
      segarkanPerGrup();
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
      var ke = el("button", "tombol-sekunder lebar", "Lihat grup berikutnya (Hitung › Per grup)");
      ke.type = "button";
      ke.addEventListener("click", function () {
        tampil("kartu-simpan", false);
        tampil("kartu-langkah", true);
        pindahTab("hitung", "pergrup");
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
    segarkanPerGrup();
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

  var nomorMuatPeriode = 0;   // hanya balasan terbaru yang digambar (tab bisa memicu beberapa muat berurutan)
  function muatPeriode(tanggal) {
    var nomor = ++nomorMuatPeriode;
    if (!statusP) teks("periode-info", "Memuat…");
    return panggil("statusPeriode", { tanggal: tanggal || "" }).then(function (d) {
      if (nomor !== nomorMuatPeriode) return;
      statusP = d;
      muatAcuan();
      gambarPilihPeriode(d);
      gambarHeader(d);
      gambarAlur(d);
      gambarBerikut(d);
      gambarGrup(d);
      gambarSyarat(d);
      gambarKonsolidasi(d);
      gambarVersi(d);
      gambarPengaturan();
      sinkronBatch(d);
      document.querySelectorAll("#rw-tabel tr[data-tanggal]").forEach(function (tr) {
        tr.classList.toggle("terpilih", tr.dataset.tanggal === d.tanggal);
      });
    }).catch(function (e) {
      if (nomor !== nomorMuatPeriode) return;
      teks("periode-info", e.message);
      teks("hp-sub", e.message);
    });
  }

  // Pemilih periode global di header (Tahap 5f)
  function gambarPilihPeriode(d) {
    var sel = $("pilih-periode");
    sel.innerHTML = "";
    var ada = {};
    (d.daftarPeriode || []).forEach(function (p) { ada[p.tanggal] = p; });
    if (d.tanggalMaster && !ada[d.tanggalMaster]) ada[d.tanggalMaster] = { tanggal: d.tanggalMaster, status: "Terbuka", jumlahGrup: 0 };
    if (d.tanggal && !ada[d.tanggal]) ada[d.tanggal] = { tanggal: d.tanggal, status: d.status, jumlahGrup: 0 };
    Object.keys(ada).sort().reverse().forEach(function (t) {
      var p = ada[t];
      var o = el("option", "", tanggalPanjang(t) + (t === d.tanggalMaster ? " · Master" : "") +
                               (p.status === "Final" ? " · Final" : "") + " · " + p.jumlahGrup + " grup");
      o.value = t;
      if (t === d.tanggal) o.selected = true;
      sel.appendChild(o);
    });
  }
  $("pilih-periode").addEventListener("change", function () {
    var t = this.value;
    muatPeriode(t);
    // Analisis › Komposisi mengikuti periode header bila periode itu punya data
    if (dataPj && (dataPj.daftarPeriode || []).indexOf(t) >= 0 && (dataPj.kini || {}).tanggal !== t) muatPenjelasan(t);
  });

  // Periode dari tempat lain (mis. baris tabel Analisis › Tren) → pemilih periode header ikut pindah
  function pilihPeriodeGlobal(t) {
    if (!t || (statusP && statusP.tanggal === t)) return;
    muatPeriode(t);
  }

  function gambarHeader(d) {
    var final = d.status === "Final";
    var bagian = [];
    if (d.periodeMaster) bagian.push(d.kodeKCMaster ? "Workbook " + d.kodeKCMaster : "Workbook: belum ada KC dicentang");
    else bagian.push("Hanya-baca · Master!C4 = " + (d.tanggalMaster ? tanggalPanjang(d.tanggalMaster) : "kosong"));
    bagian.push(d.metode ? NAMA_METODE[d.metode] : "metode belum ditetapkan");
    bagian.push(d.grupTersimpan + "/" + d.jumlahGrup + " grup tersimpan");
    teks("hp-sub", bagian.join(" · "));
    var chip = $("hp-status");
    chip.hidden = false;
    chip.className = "chip-status " + (final ? "final" : "terbuka");
    chip.textContent = final ? "Final" : "Terbuka";
    teks("rs-tanggal", tanggalPanjang(d.tanggal));
  }

  var NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  function tanggalPanjang(t) { return t ? (+t.substr(8, 2)) + " " + NAMA_BULAN[+t.substr(5, 2) - 1] + " " + t.substr(0, 4) : ""; }

  // =====================================================================
  // Tahap 4e: periode Master selalu terbaca (Hitung › Per grup)
  //   - header menampilkan Master!C4 dan tombol "Periksa ulang Master";
  //   - saat user kembali ke panel (fokus) setelah mengubah Master di Excel, periode & centang KC
  //     dicek; bila berubah, seluruh panel dan Hitung › Manual dimuat ulang otomatis.
  // =====================================================================
  var sedangCekMaster = false, cekTerakhir = 0;
  function cekMaster(paksa) {
    if (sedangCekMaster || run || batchAktif) return;
    if (!paksa && Date.now() - cekTerakhir < 1500) return;
    sedangCekMaster = true;
    cekTerakhir = Date.now();
    panggil("infoMaster").then(function (m) {
      if (!m.ditemukan) { teks("pg-master", "Workbook aplikasi CKPN tidak terbuka"); return; }
      var berubah = !statusP || m.tanggal !== statusP.tanggalMaster || m.kodeKC !== statusP.kodeKCMaster;
      if (paksa || berubah) {
        muatPeriode();     // "" = ikuti Master!C4; header, Ringkasan, Hitung, Tetapkan ikut pindah
        muatPersiapan();   // Hitung › Manual ikut segar
      }
    }).catch(function () {}).then(function () { sedangCekMaster = false; });
  }
  $("btn-periksa-master").addEventListener("click", function () { cekMaster(true); });
  window.addEventListener("focus", function () { cekMaster(false); });
  document.addEventListener("visibilitychange", function () { if (!document.hidden) cekMaster(false); });

  function gambarAlur(d) {
    teks("pg-master", d.tanggalMaster ? tanggalPanjang(d.tanggalMaster) : "Master!C4 belum diisi");
    $("pg-master-ket").textContent = d.periodeMaster ? "" :
      "Header menampilkan " + tanggalPanjang(d.tanggal) + " (hanya-baca). Hitung & simpan selalu memakai periode Master!C4 — " +
      "pilih periode Master di header atau ubah Master!C4 untuk menghitung periode lain.";
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
      ? "Metode konsolidasi " + d.tahun + ": " + NAMA_METODE[d.metode] + (d.metodeTahun && d.metodeTahun !== d.tahun ? " (mengikuti " + d.metodeTahun + ")" : "") +
        ((d.grup || []).some(function (g) { return !g.topN; }) ? " · Top-N sebagian grup belum ditetapkan" : "")
      : "Metode konsolidasi belum ditetapkan — buka ⚙ Pengaturan › Susunan grup.");
    if (d.infoPengirim && d.bolehMenulis) baris.push(d.infoPengirim);
    if (!d.bolehMenulis) baris.push(d.infoPengirim);
    baris.forEach(function (b) { info.appendChild(el("div", "", b)); });
  }

  // ---------------- Hitung › Per grup: baris grup ringkas yang bisa dibuka (Tahap 5f) ----------------
  var grupTerbuka = {};   // kodeKC → true/false (pilihan user dipertahankan saat daftar digambar ulang)
  var LABEL_STATUS_SINGKAT = { "belum": "Belum", "dihitung": "Dihitung", "tersimpan": "Tersimpan", "dihitung-ulang": "Dihitung ulang" };

  function gambarGrup(d) {
    var w = $("grup-isi");
    w.innerHTML = "";
    var final = d.status === "Final";
    var bolehHitung = d.periodeMaster && !final;
    var bolehSemua = bolehHitung && d.bolehMenulis && d.susunanDitetapkan;
    $("btn-batch-opsi").disabled = !bolehSemua;
    $("btn-batch-opsi").title = bolehSemua ? "" : "Tersedia untuk periode Master yang belum Final, setelah susunan grup ditetapkan, bagi user pengirim.";
    if (!bolehSemua) tampil("kartu-batch-opsi", false);

    var grup = d.grup || [];
    var nWb = grup.filter(function (g) { return g.diWorkbook; }).length;
    var nBelumSimpan = grup.filter(function (g) { return g.status === "dihitung" || g.status === "dihitung-ulang"; }).length;
    $("grup-ringkas").textContent = grup.length + " grup · " + d.grupTersimpan + " tersimpan" +
      (nWb ? " · " + nWb + " di workbook" : "") + (nBelumSimpan ? " · " + nBelumSimpan + " belum disimpan" : "") +
      (final ? " · periode Final" : !d.periodeMaster ? " · hanya-baca" : "");
    var kunciTotal = d.metode === "mig" ? "mig_total" : "nf_total";

    grup.forEach(function (g) {
      var det = el("details", "baris-grup" + (g.diWorkbook ? " di-workbook" : ""));
      var bawaanBuka = g.diWorkbook || g.status === "dihitung" || g.status === "dihitung-ulang";
      det.open = grupTerbuka[g.kodeKC] !== undefined ? grupTerbuka[g.kodeKC] : bawaanBuka;
      det.addEventListener("toggle", function () { grupTerbuka[g.kodeKC] = det.open; });

      var sum = el("summary");
      var kiri = el("span", "bg-kiri");
      kiri.appendChild(el("span", "nama-grup", g.nama));
      kiri.appendChild(el("span", "teks-kecil", g.kc.join(", ") + (g.topN ? " · Top-N " + g.topN : " · Top-N belum ditetapkan") +
        (g.diWorkbook ? " · di workbook" : "")));
      sum.appendChild(kiri);
      var kanan = el("span", "bg-kanan");
      var rg0 = g.aktif ? (g.aktif.ringkasan || {}) : null;
      kanan.appendChild(el("span", "bg-angka", rg0 ? jt(rg0[kunciTotal]) : "—"));
      var chip = el("span", "status-grup " + g.status,
        g.status === "tersimpan" ? "v" + g.aktif.versi : LABEL_STATUS_SINGKAT[g.status] + (g.aktif ? " · v" + g.aktif.versi : ""));
      chip.title = LABEL_STATUS[g.status] + (g.aktif ? " · versi aktif v" + g.aktif.versi : "");
      kanan.appendChild(chip);
      sum.appendChild(kanan);
      det.appendChild(sum);

      var isi = el("div", "bg-isi");
      isi.appendChild(el("div", "teks-kecil", LABEL_STATUS[g.status] + (g.aktif ? " v" + g.aktif.versi : "")));
      if (g.aktif) {
        var rg = g.aktif.ringkasan || {};
        var angka = el("div", "angka-grup");
        [["Net Flow" + (d.metode === "nf" ? " ★" : ""), rg.nf_total], ["Migration" + (d.metode === "mig" ? " ★" : ""), rg.mig_total],
         ["PPKA", g.aktif.ppkaGrup]].forEach(function (x) {
          var c = el("div", "", x[0]); c.appendChild(el("b", "", rp(x[1]))); angka.appendChild(c);
        });
        isi.appendChild(angka);
        isi.appendChild(el("div", "teks-kecil", "Disimpan " + g.aktif.waktu + " oleh " + g.aktif.pengguna));
      } else if (g.dihitungDiPC) {
        isi.appendChild(el("div", "teks-kecil", "Dihitung " + g.waktuHitung + " di PC ini"));
      }

      // ---- aksi (sama seperti sebelumnya) ----
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
      if (g.aktif) {
        var bData = tombol("Lihat / koreksi data", "tombol-sekunder", function () { bukaKoreksi(g.aktif.id, g.nama); });
        bData.title = "Tabel CKPN Individu & LGD CS versi aktif — koreksi tanpa proses ulang";
        aksi.appendChild(bData);
      }
      if (g.versi && g.versi.length) {
        var bv = tombol("Riwayat versi (" + g.versi.length + ")", "tautan", function () { pindahTab("analisis", "versi"); });
        aksi.appendChild(bv);
      }
      if (aksi.childNodes.length) isi.appendChild(aksi);
      det.appendChild(isi);
      w.appendChild(det);
    });

    if ((d.runLuarSusunan || []).length)
      w.appendChild(el("div", "teks-kecil jarak-atas", (d.runLuarSusunan.length) +
        " kiriman di luar susunan grup (tidak ikut konsolidasi) — lihat Analisis › Versi."));
  }

  // ---------------- Ringkasan: langkah berikutnya (Tahap 5f) ----------------
  function gambarBerikut(d) {
    var w = $("ringkas-berikut");
    w.innerHTML = "";
    var final = d.status === "Final";
    var grup = d.grup || [];
    var belumSimpan = grup.filter(function (g) { return g.status === "dihitung" || g.status === "dihitung-ulang"; });
    var belum = grup.filter(function (g) { return !g.aktif; });
    var pesan = "", cta = "", aksi = null;
    function nama(list) { return list.map(function (g) { return g.nama; }).join(", "); }

    if (final) {
      pesan = "Periode sudah ditetapkan (Final) oleh " + d.dikunciOleh + ". Konsolidasi dan usulan jurnal tersimpan.";
      cta = "Lihat Tetapkan →"; aksi = function () { pindahTab("tetapkan"); };
    } else if (!d.susunanDitetapkan || !d.metode) {
      pesan = "Tetapkan susunan grup, metode konsolidasi, dan Top-N tahun " + d.tahun + ".";
      cta = "Buka Pengaturan →"; aksi = function () { bukaSusunan(); };
    } else if (!d.periodeMaster && !d.siapKonsolidasi) {
      pesan = "Periode ini belum lengkap (" + d.grupTersimpan + "/" + d.jumlahGrup + " grup tersimpan). Untuk menghitungnya, " +
        "ubah Master!C4 ke " + tanggalPanjang(d.tanggal) + ".";
    } else if (belumSimpan.length && d.periodeMaster) {
      pesan = "Review & simpan hasil hitung di workbook: " + nama(belumSimpan) + ".";
      cta = "Hitung →"; aksi = function () { pindahTab("hitung", "pergrup"); };
    } else if (belum.length) {
      pesan = "Hitung & simpan " + belum.length + " grup lagi: " + nama(belum) + ".";
      cta = "Hitung →"; aksi = function () { pindahTab("hitung", "pergrup"); };
    } else if (d.siapKonsolidasi) {
      pesan = "Semua grup tersimpan. Cek konsolidasi vs PPKA dan usulan jurnal, lalu tetapkan periode.";
      cta = "Tetapkan →"; aksi = function () { pindahTab("tetapkan"); };
    }
    w.hidden = !pesan;
    if (!pesan) return;
    var kiri = el("div");
    kiri.appendChild(el("div", "label-kecil", "Berikutnya"));
    kiri.appendChild(el("div", "", pesan));
    w.appendChild(kiri);
    if (cta) w.appendChild(tombol(cta, "tombol-sekunder", aksi));
  }

  // ---------------- Tetapkan: syarat penetapan (Tahap 5f) ----------------
  function gambarSyarat(d) {
    var ul = $("syarat-isi");
    ul.innerHTML = "";
    var final = d.status === "Final";
    var grup = d.grup || [];
    function item(status, label, rinci) {
      var li = el("li", "syarat s-" + status);
      li.appendChild(el("span", "syarat-ikon", status === "ok" ? "✓" : status === "awas" ? "!" : "✕"));
      var t = el("div");
      t.appendChild(el("div", "", label));
      if (rinci) t.appendChild(el("div", "teks-kecil", rinci));
      li.appendChild(t);
      ul.appendChild(li);
    }
    if (final) item("ok", "Periode Final", "Dikunci oleh " + d.dikunciOleh + " · " + d.dikunciWaktu);
    item(d.susunanDitetapkan ? "ok" : "belum", "Susunan grup tahun " + d.tahun + " ditetapkan",
      d.susunanDitetapkan ? (d.susunanTahun !== d.tahun ? "Mengikuti susunan " + d.susunanTahun : (d.susunanDasar || "")) : "Atur di ⚙ Pengaturan.");
    item(d.metode ? "ok" : "belum", "Metode konsolidasi ditetapkan", d.metode ? NAMA_METODE[d.metode] + " + LGD weighted" : "Atur di ⚙ Pengaturan.");
    var tanpaTopN = grup.filter(function (g) { return !g.topN; });
    item(tanpaTopN.length ? "awas" : "ok", "Top-N CKPN Individu per grup",
      tanpaTopN.length ? "Belum ditetapkan: " + tanpaTopN.map(function (g) { return g.nama; }).join(", ") : "");
    item(d.siapKonsolidasi ? "ok" : "belum", d.grupTersimpan + "/" + d.jumlahGrup + " grup tersimpan",
      d.siapKonsolidasi ? "" : "Simpan grup yang belum: " + grup.filter(function (g) { return !g.aktif; }).map(function (g) { return g.nama; }).join(", "));
    var belumSimpan = grup.filter(function (g) { return g.status === "dihitung" || g.status === "dihitung-ulang"; });
    if (belumSimpan.length)
      item("awas", "Hasil hitung di workbook belum disimpan: " + belumSimpan.map(function (g) { return g.nama; }).join(", "),
        "Konsolidasi memakai versi yang tersimpan, bukan isi sheet saat ini.");
    if ((d.runLuarSusunan || []).length)
      item("awas", d.runLuarSusunan.length + " kiriman di luar susunan grup", "Tidak ikut konsolidasi. Lihat Analisis › Versi.");
    if (!final) item(d.bolehMenulis ? "ok" : "belum", "Hak menetapkan", d.bolehMenulis ? "" : (d.infoPengirim || "User ini hanya bisa melihat."));
  }

  // ---------------- Analisis › Versi: riwayat versi semua grup (Tahap 5f) ----------------
  function gambarVersi(d) {
    var w = $("versi-isi");
    w.innerHTML = "";
    var final = d.status === "Final";
    teks("versi-judul", "Riwayat versi · " + tanggalPanjang(d.tanggal));
    $("versi-sub").textContent = "Pilih periode lain di header. Data = lihat isi versi; xlsx = snapshot saat disimpan." +
      (final ? " Periode Final: versi tidak bisa dihapus." : "");
    var ada = false;
    (d.grup || []).forEach(function (g) {
      var blok = el("div", "blok-versi");
      blok.appendChild(el("div", "nama-grup", g.nama));
      blok.appendChild(el("div", "teks-kecil", g.kc.join(", ") + " · " + LABEL_STATUS[g.status]));
      if (g.versi && g.versi.length) {
        ada = true;
        var det = daftarVersi(g.versi, d, final);
        det.open = true;
        det.querySelector("summary").textContent = g.versi.length + " versi";
        blok.appendChild(det);
      } else blok.appendChild(el("div", "teks-kecil", "Belum ada versi tersimpan."));
      w.appendChild(blok);
    });
    if ((d.runLuarSusunan || []).length) {
      var luar = el("div", "blok-versi");
      luar.appendChild(el("div", "nama-grup", "Di luar susunan grup"));
      luar.appendChild(el("div", "teks-kecil", "Kiriman dengan kombinasi KC yang tidak ada di susunan — tidak ikut konsolidasi."));
      luar.appendChild(daftarVersi(d.runLuarSusunan, d, final, true));
      w.appendChild(luar);
      ada = true;
    }
    if (!ada && !(d.grup || []).length) w.textContent = "Belum ada grup.";
  }

  // ---------------- ⚙ Pengaturan: ringkasan nilai (Tahap 5f) ----------------
  var infoPing = null;
  function gambarPengaturan() {
    var d = statusP;
    if (d) {
      var grup = d.grup || [];
      teks("pt-tahun-judul", "Tahun " + d.tahun);
      teks("pt-susunan", d.susunanDitetapkan
        ? grup.length + " grup" + (d.susunanTahun !== d.tahun ? " · mengikuti " + d.susunanTahun : "") + " · " +
          grup.map(function (g) { return g.nama; }).join(", ")
        : "Belum ditetapkan");
      teks("pt-metode", d.metode ? NAMA_METODE[d.metode] + " + LGD weighted" +
        (d.metodeTahun && d.metodeTahun !== d.tahun ? " · mengikuti " + d.metodeTahun : "") : "Belum ditetapkan");
      teks("pt-topn", grup.map(function (g) { return g.nama + " " + (g.topN || "—"); }).join(" · "));
      teks("pt-dasar", d.susunanDasar || "Belum diisi");
      var adaFinal = (d.daftarPeriode || []).some(function (p) { return p.status === "Final" && p.tanggal.substr(0, 4) === String(d.tahun); });
      var k = $("pt-kunci");
      k.className = adaFinal ? "peringatan-box" : "teks-kecil";
      k.textContent = adaFinal
        ? "Ada periode Final di tahun " + d.tahun + ": metode konsolidasi dan Top-N terkunci sampai kunci periode tersebut dibuka."
        : "Metode konsolidasi dan Top-N terkunci setelah ada periode Final di tahun ini.";
    }
    teks("pt-pdlgd", !infoAcuan ? "—" : !infoAcuan.tahunan ? "Dihitung setiap bulan" :
      "Setahun sekali · " + (infoAcuan.desember || !infoAcuan.tahunAcuan ? "Desember dihitung penuh" : "acuan Desember " + infoAcuan.tahunAcuan));
    var boleh = infoPing && infoPing.bolehMenulis !== undefined ? infoPing.bolehMenulis : d ? d.bolehMenulis : null;
    teks("pt-pengirim", boleh === null ? "—" : boleh ? "Anda: lihat & simpan" : "Anda: hanya lihat");
    $("pt-pengirim-user").textContent = infoPing
      ? "User Windows Anda: " + infoPing.user + " — " + (boleh ? "boleh menyimpan." : "belum terdaftar; tambahkan nama ini ke pengirim.txt agar bisa menyimpan.")
      : "";
  }
  $("btn-atur-pengirim").addEventListener("click", function () {
    var buka = $("pt-pengirim-info").hidden;
    tampil("pt-pengirim-info", buka);
    this.setAttribute("aria-expanded", buka ? "true" : "false");
  });

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
      if (!v.dihapus) aksi.appendChild(tombol("Data", "", function () { bukaKoreksi(v.id, v.kodeKC); }));
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
    var pakai = pakaiAcuan();
    var pesan = (untukEdit ? "Buka " + g.nama + " untuk diedit.\n\n" : "Hitung " + g.nama + ".\n\n") +
      "Centang KC di Master akan diubah menjadi: " + g.kc.join(", ") + "\n" +
      (pakai ? teksAcuanGrup(g.kodeKC) + "\n" : "Semua langkah (Individu s.d. Summary) dijalankan dengan penyesuaian tersimpan.\n") +
      "Isi sheet hasil saat ini akan ditimpa. Lanjutkan?";
    if (!window.confirm(pesan)) return;
    panggil("hitungGrup", { kodeKC: g.kodeKC, terapkanPenyesuaian: true, pakaiAcuan: pakai }).then(function () {
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
      var t = tombol("Atur di ⚙ Pengaturan", "tautan", function () { bukaSusunan("susunan-metode"); });
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

  // ---------------- Susunan grup (editor di ⚙ Pengaturan) ----------------
  function bukaSusunan(fokus) {
    if (tabAktif !== "pengaturan") pindahTab("pengaturan");
    if (!statusP) {
      muatPeriode().then(function () { if (statusP) bukaSusunan(fokus); });
      return;
    }
    isiEditorSusunan();
    var target = fokus === "susunan-metode" ? document.querySelector('input[name="susunan-metode"]').closest("fieldset")
               : fokus ? $(fokus) : null;
    (target || $("kartu-susunan")).scrollIntoView({ block: "start" });
    if (fokus === "susunan-dasar") $("susunan-dasar").focus();
  }
  $("btn-atur-susunan").addEventListener("click", function () { bukaSusunan(); });
  document.querySelectorAll('[data-atur="susunan"]').forEach(function (b) {
    b.addEventListener("click", function () { bukaSusunan(b.getAttribute("data-fokus")); });
  });

  function isiEditorSusunan() {
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
      inp.addEventListener("input", gambarTopNGrup);
      w.appendChild(lab); w.appendChild(inp);
    });
    // nilai awal Top-N per nama grup: ketetapan tersimpan → Master!C10 → 10
    $("susunan-topn-grup").innerHTML = "";   // buang isian sesi sebelumnya
    topNIsian = {};
    (statusP.grup || []).forEach(function (g) { if (g.topN) topNIsian[g.nama] = g.topN; });
    gambarTopNGrup();
    $("susunan-dasar").value = statusP.susunanDasar || "";
    document.querySelectorAll('input[name="susunan-metode"]').forEach(function (r) {
      r.checked = r.value === statusP.metode;
    });
    $("susunan-hasil").innerHTML = "";
    $("btn-susunan-simpan").disabled = !statusP.bolehMenulis;
    tampil("kartu-susunan", true);
  }
  $("btn-susunan-perkc").addEventListener("click", function () {
    ["KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100"].forEach(function (kc) { $("sg-" + kc).value = kc; });
    gambarTopNGrup();
  });

  // ---- Top-N per grup: daftar isian dibangun ulang setiap nama grup berubah ----
  var topNIsian = {};   // nama grup → Top-N yang sudah diketik (dipertahankan saat daftar dibangun ulang)
  function namaGrupSusunan() {
    var urut = [];
    ["KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100"].forEach(function (kc) {
      var n = $("sg-" + kc).value.trim();
      if (n && urut.indexOf(n) < 0) urut.push(n);
    });
    return urut;
  }
  function gambarTopNGrup() {
    var w = $("susunan-topn-grup");
    w.querySelectorAll("input").forEach(function (i) { if (i.value !== "") topNIsian[i.dataset.nama] = Number(i.value); });
    w.innerHTML = "";
    var bawaan = infoTopN ? infoTopN.topNMaster : 10;
    namaGrupSusunan().forEach(function (n, idx) {
      var lab = el("label", "kode", n);
      lab.setAttribute("for", "tn-" + idx);
      var inp = el("input", "isian angka lebar-kecil");
      inp.type = "number"; inp.min = "1"; inp.max = "1000"; inp.step = "1"; inp.id = "tn-" + idx;
      inp.dataset.nama = n;
      inp.value = topNIsian[n] || bawaan;
      w.appendChild(lab); w.appendChild(inp);
    });
  }
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
    var topN = {}, salah = [], diganti = [];
    $("susunan-topn-grup").querySelectorAll("input").forEach(function (i) {
      var n = Number(i.value);
      if (!(n >= 1 && n <= 1000 && Math.floor(n) === n)) salah.push(i.dataset.nama);
      topN[i.dataset.nama] = n;
    });
    if (salah.length) {
      $("susunan-hasil").innerHTML = "";
      $("susunan-hasil").appendChild(el("div", "peringatan-box", "Isi Top-N (bilangan bulat 1–1000) untuk grup: " + salah.join(", ") + "."));
      return;
    }
    (statusP.grup || []).forEach(function (g) {
      if (g.topN && topN[g.nama] !== undefined && topN[g.nama] !== g.topN) diganti.push(g.nama + " " + g.topN + " → " + topN[g.nama]);
    });
    if (diganti.length &&
        !window.confirm("Top-N diganti: " + diganti.join(", ") + ". Perubahan ini tercatat di log aktivitas. Lanjutkan?")) return;
    if (statusP.metode && m.value !== statusP.metode &&
        !window.confirm("Metode konsolidasi diganti dari " + NAMA_METODE[statusP.metode] + " ke " + NAMA_METODE[m.value] +
                        ". Perubahan ini tercatat di log aktivitas. Lanjutkan?")) return;
    panggil("simpanSusunan", {
      tahun: statusP.tahun, peta: peta, dasar: $("susunan-dasar").value,
      metode: m.value, kebijakanSaldo: "ckpn", topN: topN
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
  // 8b. Hitung semua grup sekaligus (Tahap 4d)
  // =====================================================================
  var batchAktif = false, runTerakhirBatch = null;
  var batchInfo = null;   // { tanggal, versiAwal: {kodeKC: versi aktif saat batch dimulai}, berhenti }
  var LABEL_BATCH = {
    "antre": "menunggu", "menghitung": "menghitung…", "tersimpan": "tersimpan", "tersimpan-temuan": "tersimpan · cek temuan", "perlu-review": "perlu review",
    "gagal": "gagal", "dibatalkan": "dibatalkan"
  };

  $("btn-batch-opsi").addEventListener("click", function () {
    tampil("kartu-batch-opsi", !$("kartu-batch-opsi").hidden ? false : true);
  });
  $("btn-batch-batal").addEventListener("click", function () { tampil("kartu-batch-opsi", false); });
  $("btn-batch-mulai").addEventListener("click", function () {
    var b = this;
    var berhenti = document.querySelector('input[name="batch-temuan"]:checked').value === "berhenti";
    if (!window.confirm("Hitung semua grup periode " + (statusP ? statusP.tanggal : "") + "?\n\n" +
        "Centang KC dan Top-N di Master akan diubah untuk setiap grup, dan isi sheet hasil ditimpa. " +
        (berhenti ? "Grup tanpa temuan langsung disimpan; proses berhenti di grup pertama yang punya temuan."
                  : "SEMUA grup langsung disimpan sebagai versi baru, termasuk yang punya temuan (cek & koreksi sesudahnya).") +
        (pakaiAcuan() ? "\n\nMetode PD & LGD setahun sekali: " + ringkasAcuan() : "")))
      return;
    b.disabled = true;
    panggil("hitungSemuaGrup", { lewatiTersimpan: $("batch-lewati").checked, berhentiBilaTemuan: berhenti, pakaiAcuan: pakaiAcuan() })
      .then(function () { tampil("kartu-batch-opsi", false); window.scrollTo(0, 0); })
      .catch(function (e) { alert(e.message); })
      .then(function () { b.disabled = false; });
  });
  $("btn-batch-tutup").addEventListener("click", function () { tampil("kartu-batch", false); });

  onEvent("batchMulai", function (d) {
    batchAktif = true;
    batchInfo = { tanggal: d.tanggal, versiAwal: {}, berhenti: d.berhentiBilaTemuan };
    ((statusP && statusP.tanggal === d.tanggal && statusP.grup) || []).forEach(function (g) {
      batchInfo.versiAwal[g.kodeKC] = g.aktif ? g.aktif.versi : 0;
    });
    runTerakhirBatch = null;
    tampil("kartu-hasil", false);
    tampil("kartu-simpan", false);
    tampil("btn-batch-tutup", false);
    teks("batch-judul", "Hitung semua grup · " + d.tanggal);
    teks("batch-sub", d.grup.length + " grup · " + (d.berhentiBilaTemuan ? "review dulu bila ada temuan" : "simpan semua otomatis") +
      (d.pakaiAcuan ? " · PD & LGD acuan Desember" : ""));
    var ol = $("batch-daftar");
    ol.innerHTML = "";
    d.grup.forEach(function (g, i) {
      var li = el("li", "antre");
      li.id = "bg-" + i;
      li.dataset.kode = g.kodeKC;
      li.dataset.nama = g.nama;
      li.dataset.status = "antre";
      var kepala = el("div", "baris-antara");
      var kiri = el("div");
      kiri.appendChild(el("div", "nama-grup", g.nama));
      kiri.appendChild(el("div", "teks-kecil", g.kodeKC + (g.topN ? " · Top-N " + g.topN : "")));
      kepala.appendChild(kiri);
      kepala.appendChild(el("span", "status-grup", LABEL_BATCH.antre));
      li.appendChild(kepala);
      li.appendChild(el("ul", "batch-temuan"));
      ol.appendChild(li);
    });
    $("batch-ringkas").innerHTML = "";
    tampil("kartu-batch", true);
  });

  onEvent("batchGrup", function (d) {
    var li = $("bg-" + (d.ke - 1));
    if (!li) return;
    li.className = d.status;
    li.dataset.status = d.status;
    var chip = li.querySelector(".status-grup");
    chip.className = "status-grup " + (d.status === "tersimpan" ? "tersimpan" :
      d.status === "perlu-review" || d.status === "menghitung" || d.status === "tersimpan-temuan" ? "dihitung" : "");
    chip.textContent = LABEL_BATCH[d.status] + (d.versi ? " v" + d.versi : "");
    var ul = li.querySelector(".batch-temuan");
    ul.innerHTML = "";
    (d.temuan || []).forEach(function (t) { ul.appendChild(el("li", "", t)); });
    if (d.acuan && d.status !== "menghitung") ul.appendChild(el("li", "ok-ringan", d.acuan));
    if ((d.status === "tersimpan" || d.status === "tersimpan-temuan") && (d.nfTotal !== undefined || d.migTotal !== undefined))
      ul.appendChild(el("li", "ok-ringan", "Net Flow " + rp(d.nfTotal) + " · Migration " + rp(d.migTotal)));
    if (d.status === "tersimpan-temuan" && d.runId) {
      var aksi = el("div", "aksi-grup aksi-batch-data");
      aksi.appendChild(tombol("Lihat / koreksi data", "tombol-sekunder", function () { bukaKoreksiDariBatch(d.runId, d.nama); }));
      li.appendChild(aksi);
    }
  });

  onEvent("batchSelesai", function (d) {
    batchAktif = false;
    if (batchInfo) batchInfo.selesai = d;
    var n = { tersimpan: 0, temuan: 0, review: 0, gagal: 0 };
    (d.hasil || []).forEach(function (h) {
      if (h.status === "tersimpan" || h.status === "tersimpan-temuan") n.tersimpan++;
      if (h.status === "tersimpan-temuan") n.temuan++;
      if (h.status === "perlu-review") n.review++;
      if (h.status === "gagal") n.gagal++;
    });
    var w = $("batch-ringkas");
    w.innerHTML = "";
    var teksRingkas = n.tersimpan + " grup tersimpan otomatis" + (n.temuan ? " (" + n.temuan + " dengan temuan)" : "") +
      (n.review ? " · " + n.review + " perlu review" : "") + (n.gagal ? " · " + n.gagal + " gagal" : "") +
      (d.sisa ? " · " + d.sisa + " grup belum dihitung" : "");
    tambahTombolBaris();
    if (d.status === "selesai") w.appendChild(el("div", "status-hasil " + (n.gagal ? "batal" : "ok"), "Selesai · " + teksRingkas));
    if (n.temuan)
      w.appendChild(el("div", "teks-kecil", "Grup bertanda \"cek temuan\" sudah tersimpan. Cek kontrak/debitur yang disebut di temuannya lewat " +
        "Lihat / koreksi data; bila nilai agunan, biaya, atau baris LGD CS perlu diubah, simpan sebagai versi baru di sana. " +
        "Agar berlaku juga untuk bulan berikutnya, isi pula Data penyesuaian."));
    else if (d.status === "perlu-review") {
      w.appendChild(el("div", "status-hasil batal", "Berhenti untuk review · " + teksRingkas));
      w.appendChild(el("div", "teks-kecil", "Sheet berisi grup yang perlu direview. Cek temuan di bawah, edit bila perlu, lalu simpan. " +
        "Setelah itu jalankan lagi Hitung semua grup (grup tersimpan dilewati) untuk melanjutkan."));
    } else if (d.status === "dibatalkan") w.appendChild(el("div", "status-hasil batal", "Dibatalkan · " + teksRingkas));
    else w.appendChild(el("div", "status-hasil err", "Gagal: " + (d.error || "") + " · " + teksRingkas));
    if (n.review && d.status !== "perlu-review")
      w.appendChild(el("div", "teks-kecil", "Langkah berikutnya: kerjakan grup bertanda satu per satu — tekan Hitung grup ini pada barisnya, " +
        "review, lalu Simpan. Status di kartu ini ikut diperbarui. Tidak perlu menjalankan Hitung semua grup lagi."));
    tampil("btn-batch-tutup", true);
    // Grup tempat proses berhenti: tampilkan kartu hasil + review agar bisa langsung disimpan
    if (d.status === "perlu-review" && runTerakhirBatch) { tampilkanHasil(runTerakhirBatch); muatReview(); }
    else tampil("kartu-langkah", true);
    muatPersiapan();
    segarkanPerGrup();
  });

  // Grup tersimpan dengan temuan → buka editor koreksi di Hitung › Per grup
  function bukaKoreksiDariBatch(runId, nama) {
    if (tabAktif !== "hitung" || subAktif.hitung !== "pergrup") pindahTab("hitung", "pergrup");
    bukaKoreksi(runId, nama);
  }

  // Setelah batch selesai: baris yang belum tersimpan diberi tombol "Hitung grup ini"
  function tambahTombolBaris() {
    document.querySelectorAll("#batch-daftar > li").forEach(function (li) {
      var st = li.dataset.status;
      if (st === "tersimpan" || st === "tersimpan-temuan" || st === "tersimpan-manual" || li.querySelector(".aksi-batch")) return;
      var aksi = el("div", "aksi-grup aksi-batch");
      aksi.appendChild(tombol("Hitung grup ini", "tombol-sekunder", function () {
        hitungGrup({ nama: li.dataset.nama, kodeKC: li.dataset.kode, kc: li.dataset.kode.split(",") });
      }));
      li.appendChild(aksi);
    });
  }

  // Selaraskan kartu "Hitung semua grup" dengan status grup terbaru (dipanggil tiap status periode dimuat):
  // grup bertanda yang kemudian dihitung/disimpan satu per satu ikut berubah statusnya.
  function sinkronBatch(d) {
    if (!batchInfo || batchAktif || $("kartu-batch").hidden || d.tanggal !== batchInfo.tanggal) return;
    var sisaReview = 0, total = 0;
    document.querySelectorAll("#batch-daftar > li").forEach(function (li) {
      total++;
      var g = null;
      (d.grup || []).forEach(function (x) { if (x.kodeKC === li.dataset.kode) g = x; });
      var st = li.dataset.status;
      var simpan = st === "tersimpan" || st === "tersimpan-temuan" || st === "tersimpan-manual";
      if (!g) { if (!simpan) sisaReview++; return; }
      if (!simpan) {
        var awal = batchInfo.versiAwal[li.dataset.kode] || 0;
        var chip = li.querySelector(".status-grup");
        if (g.aktif && g.aktif.versi > awal && (g.status === "tersimpan")) {
          li.dataset.status = "tersimpan-manual";
          li.className = "tersimpan";
          chip.className = "status-grup tersimpan";
          chip.textContent = "tersimpan v" + g.aktif.versi + " (per grup)";
          var ul = li.querySelector(".batch-temuan");
          ul.innerHTML = "";
          ul.appendChild(el("li", "ok-ringan", "Direview & disimpan " + g.aktif.waktu + " oleh " + g.aktif.pengguna));
          var ak = li.querySelector(".aksi-batch");
          if (ak) ak.parentNode.removeChild(ak);
        } else if (g.status === "dihitung" || g.status === "dihitung-ulang") {
          chip.className = "status-grup dihitung";
          chip.textContent = "dihitung — belum disimpan";
        }
      }
      st = li.dataset.status;
      if (st !== "tersimpan" && st !== "tersimpan-temuan" && st !== "tersimpan-manual") sisaReview++;
    });
    var w = $("batch-ringkas");
    var info = w.querySelector(".sinkron-batch") || w.appendChild(el("div", "teks-kecil sinkron-batch"));
    info.className = sisaReview ? "teks-kecil sinkron-batch" : "status-hasil ok sinkron-batch";
    info.textContent = sisaReview ? "Status terkini: " + (total - sisaReview) + " dari " + total + " grup tersimpan · " + sisaReview + " belum."
                                  : "Semua " + total + " grup sudah tersimpan — lanjut ke Konsolidasi vs PPKA.";
  }

  // =====================================================================
  // 9. Tab Riwayat antarperiode (Tahap 4) — pengganti Dashboard CKPN
  // =====================================================================
  var dataRw = null;
  var fmtM = new Intl.NumberFormat("id-ID", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  var fmtM2 = new Intl.NumberFormat("id-ID", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  // Label sumbu: jumlah desimal mengikuti jarak antargaris agar tidak ada label kembar (0,3 / 0,3)
  function sumbuM(v, step) {
    var s = step / 1e9, d = s >= 1 ? 0 : s >= 0.1 ? 1 : s >= 0.01 ? 2 : 3;
    return new Intl.NumberFormat("id-ID", { minimumFractionDigits: d, maximumFractionDigits: d }).format(v / 1e9);
  }
  var BULAN_SINGKAT = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];

  function juta(v) { return v === null || v === undefined ? "—" : rp(v / 1e6); }
  function angkaM(v) { return v === null || v === undefined ? "—" : fmtM2.format(v / 1e9); }   // kolom tabel: satuan di judul
  function miliar(v) {
    if (v === null || v === undefined) return "—";
    return (Math.abs(v) < 1e10 ? fmtM2 : fmtM).format(v / 1e9) + " M";   // < 10 M: dua desimal
  }
  function labelPeriode(t) { return BULAN_SINGKAT[+t.substr(5, 2) - 1] + " " + t.substr(2, 2); }
  function pct(v) { return v === null || v === undefined || !isFinite(v) ? "—" : (v * 100).toFixed(1).replace(".", ",") + "%"; }
  function bertanda(v) { return v === null || v === undefined ? "—" : (v > 0 ? "+" : v < 0 ? "−" : "") + rp(Math.abs(v)); }

  function muatRiwayat() {
    teks("rw-info", "Memuat…");
    return panggil("riwayat").then(function (d) {
      dataRw = d;
      isiFilterRiwayat();
      gambarRiwayat();
    }).catch(function (e) { teks("rw-info", e.message); });
  }

  function isiFilterRiwayat() {
    var tahun = {}, selT = $("rw-tahun"), selG = $("rw-grup");
    var lamaT = selT.value, lamaG = selG.value;
    dataRw.periode.forEach(function (p) { tahun[p.tanggal.substr(0, 4)] = true; });
    selT.innerHTML = "";
    var o = el("option", "", "Semua"); o.value = ""; selT.appendChild(o);
    Object.keys(tahun).sort().reverse().forEach(function (t) { var x = el("option", "", t); x.value = t; selT.appendChild(x); });
    selT.value = lamaT && tahun[lamaT] ? lamaT : "";
    selG.innerHTML = "";
    var b = el("option", "", "Bank (semua grup + ABA)"); b.value = ""; selG.appendChild(b);
    dataRw.grup.forEach(function (g) { var x = el("option", "", g); x.value = g; selG.appendChild(x); });
    selG.value = dataRw.grup.indexOf(lamaG) >= 0 ? lamaG : "";
  }

  function nilaiGrup(p, nama) {
    var g = null;
    p.grup.forEach(function (x) { if (x.nama === nama && x.ada) g = x; });
    return g;
  }

  function tile(label, nilai, sub, kelasSub) {
    var t = el("div", "tile");
    t.appendChild(el("div", "tile-label", label));
    t.appendChild(el("div", "tile-nilai", nilai));
    if (sub) t.appendChild(el("div", "tile-sub" + (kelasSub ? " " + kelasSub : ""), sub));
    return t;
  }

  function gambarRiwayat() {
    if (!dataRw) return;
    var tahun = $("rw-tahun").value, grup = $("rw-grup").value;
    var list = dataRw.periode.filter(function (p) { return !tahun || p.tanggal.substr(0, 4) === tahun; });
    var tiles = $("rw-tile"), w = $("rw-grafik");
    tiles.innerHTML = ""; w.innerHTML = ""; $("rw-legenda").innerHTML = ""; $("rw-tabel").innerHTML = "";
    if (!list.length) { teks("rw-info", "Belum ada periode dengan kiriman grup tersimpan."); return; }
    teks("rw-info", list.length + " periode · " + labelPeriode(list[0].tanggal) + " s.d. " + labelPeriode(list[list.length - 1].tanggal) +
      " · hanya kiriman aktif yang sesuai susunan grup");

    var seri, akhir = list[list.length - 1];
    if (!grup) {
      teks("rw-judul-grafik", "CKPN vs PPKA — bank (termasuk ABA)");
      seri = [
        { nama: "CKPN Net Flow", singkat: "NF", kelas: "seri-1", nilai: list.map(function (p) { return p.nf; }) },
        { nama: "CKPN Migration", singkat: "Mig", kelas: "seri-2", nilai: list.map(function (p) { return p.mig; }) },
        { nama: "PPKA OJK", singkat: "PPKA", kelas: "seri-3", nilai: list.map(function (p) { return p.ppka; }) }
      ];
      var sel = akhir.selisih;
      tiles.appendChild(tile("CKPN terpakai · " + labelPeriode(akhir.tanggal), miliar(akhir.ckpn),
        akhir.metode ? NAMA_METODE[akhir.metode] + " · " + akhir.status : "metode belum ditetapkan"));
      tiles.appendChild(tile("PPKA OJK", miliar(akhir.ppka), akhir.lengkap ? "semua grup tersimpan" :
        akhir.grupTersimpan + "/" + akhir.jumlahGrup + " grup tersimpan", akhir.lengkap ? "" : "kurang"));
      tiles.appendChild(tile("Selisih CKPN − PPKA", sel === null ? "—" : (sel >= 0 ? "+" : "−") + miliar(Math.abs(sel)),
        sel === null ? "" : sel > 0 ? "PPKA < CKPN · bentuk tambahan" : sel < 0 ? "PPKA > CKPN" : "sama"));
    } else {
      teks("rw-judul-grafik", "CKPN vs PPKA — " + grup);
      function ambil(k) { return list.map(function (p) { var g = nilaiGrup(p, grup); return g ? g[k] : null; }); }
      seri = [
        { nama: "CKPN Net Flow", singkat: "NF", kelas: "seri-1", nilai: ambil("nf") },
        { nama: "CKPN Migration", singkat: "Mig", kelas: "seri-2", nilai: ambil("mig") },
        { nama: "PPKA OJK", singkat: "PPKA", kelas: "seri-3", nilai: ambil("ppka") }
      ];
      // periode terakhir yang memuat grup ini (grup bisa belum tersimpan di periode berjalan)
      var pg = null;
      list.forEach(function (p) { if (nilaiGrup(p, grup)) pg = p; });
      var ga = pg ? nilaiGrup(pg, grup) : null;
      tiles.appendChild(tile("Net Flow · " + (pg ? labelPeriode(pg.tanggal) : "—"), miliar(ga ? ga.nf : null)));
      tiles.appendChild(tile("Migration", miliar(ga ? ga.mig : null)));
      tiles.appendChild(tile("PPKA OJK", miliar(ga ? ga.ppka : null), ga ? "LGD " + pct(ga.lgd) : "belum tersimpan"));
    }

    var leg = $("rw-legenda");
    seri.forEach(function (s) {
      var i = el("span");
      i.appendChild(el("i", "kunci-garis " + s.kelas));
      i.appendChild(document.createTextNode(s.nama));
      leg.appendChild(i);
    });
    Grafik.garis(w, {
      x: list.map(function (p) { return labelPeriode(p.tanggal); }), seri: seri, tinggi: 220,
      formatSumbu: sumbuM, formatNilai: function (v) { return "Rp " + rp(v); },
      label: "Tren CKPN Net Flow, Migration, dan PPKA per periode"
    });

    // ---- tabel (tampilan setara grafik, terbaru di atas) ----
    var t = el("table", "tabel-konsolidasi tabel-rw");
    var h = el("tr");
    (grup ? ["Periode", "NF (M)", "Mig (M)", "PPKA (M)", "LGD"] : ["Periode", "NF (M)", "Mig (M)", "PPKA (M)", "Selisih"])
      .forEach(function (x) { h.appendChild(el("th", "", x)); });
    t.appendChild(h);
    list.slice().reverse().forEach(function (p) {
      // Tahap 5f: ketuk baris → pemilih periode header pindah ke periode itu
      var tr = el("tr", "baris-klik" + (statusP && statusP.tanggal === p.tanggal ? " terpilih" : ""));
      tr.dataset.tanggal = p.tanggal;
      tr.tabIndex = 0;
      tr.title = "Tampilkan " + tanggalPanjang(p.tanggal) + " di seluruh panel";
      tr.addEventListener("click", function () { pilihPeriodeGlobal(p.tanggal); });
      tr.addEventListener("keydown", function (e) {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); pilihPeriodeGlobal(p.tanggal); }
      });
      var c1 = el("td");
      c1.appendChild(el("div", "", labelPeriode(p.tanggal)));
      c1.appendChild(el("div", "teks-kecil", p.status + (p.metode ? " · " + (p.metode === "nf" ? "NF" : "Mig") : "") +
        (p.lengkap ? "" : " · " + p.grupTersimpan + "/" + p.jumlahGrup)));
      tr.appendChild(c1);
      if (grup) {
        var g = nilaiGrup(p, grup);
        [g ? g.nf : null, g ? g.mig : null, g ? g.ppka : null].forEach(function (v) { tr.appendChild(el("td", "", angkaM(v))); });
        tr.appendChild(el("td", "", g ? pct(g.lgd) : "—"));
      } else {
        tr.appendChild(el("td", p.metode === "nf" ? "dipakai" : "", angkaM(p.nf)));
        tr.appendChild(el("td", p.metode === "mig" ? "dipakai" : "", angkaM(p.mig)));
        tr.appendChild(el("td", "", angkaM(p.ppka)));
        var cs = el("td", p.selisih > 0 ? "kurang" : "", p.selisih === null ? "—" : (p.selisih >= 0 ? "+" : "−") + angkaM(Math.abs(p.selisih)));
        if (p.jurnal) cs.appendChild(el("div", "teks-kecil", "jurnal " + p.jurnal.jenis));
        tr.appendChild(cs);
      }
      t.appendChild(tr);
    });
    $("rw-tabel").appendChild(t);
    $("rw-tabel").appendChild(el("p", "teks-kecil", (grup ? "" : "Kolom yang disorot = metode terpakai. Selisih positif berarti PPKA lebih kecil dari CKPN. ") +
      "Ketuk baris periode untuk menampilkannya di seluruh panel (pemilih periode di atas ikut pindah)."));
  }

  $("rw-tahun").addEventListener("change", gambarRiwayat);
  $("rw-grup").addEventListener("change", gambarRiwayat);
  $("btn-rw-muat").addEventListener("click", muatRiwayat);

  // =====================================================================
  // 10. Tab Analisis PD (Tahap 4): sankey PD Migration & PD Net Flow
  // =====================================================================
  var dataAn = null, twAktif = null;
  var LABEL_KOL = ["Kol 1", "Kol 2", "Kol 3", "Kol 4", "Kol 5"];

  function muatSumberAnalisis() {
    return panggil("analisisSumber").then(function (d) {
      var sel = $("an-sumber"), lama = sel.value;
      sel.innerHTML = "";
      var wb = el("option", "", "Workbook saat ini (hasil hitung terakhir di sheet)"); wb.value = "wb"; sel.appendChild(wb);
      d.periode.forEach(function (p) {
        var og = document.createElement("optgroup");
        og.label = p.tanggal;
        var adaData = p.runs.some(function (r) { return r.adaData; });
        if (p.runs.length > 1) {
          var g = el("option", "", "Semua grup (dijumlah)"); g.value = "g|" + p.tanggal; g.disabled = !adaData; og.appendChild(g);
        }
        p.runs.forEach(function (r) {
          var o = el("option", "", r.nama + " (v" + r.versi + ")" + (r.adaData ? "" : " — tanpa data analisis"));
          o.value = "r|" + r.runId; o.disabled = !r.adaData; og.appendChild(o);
        });
        sel.appendChild(og);
      });
      if (lama && sel.querySelector('option[value="' + lama.replace(/"/g, "") + '"]')) sel.value = lama;
    }).catch(function (e) { teks("an-info", e.message); });
  }

  $("btn-an-muat").addEventListener("click", function () {
    var v = $("an-sumber").value || "wb", b = this;
    var args = v === "wb" ? { sumber: "workbook" } : v.charAt(0) === "g" ? { sumber: "gabungan", tanggal: v.substr(2) } :
               { sumber: "run", runId: Number(v.substr(2)) };
    b.disabled = true;
    teks("an-info", "Memuat…");
    panggil("analisisData", args, 120000).then(function (d) {
      dataAn = d; twAktif = null;
      gambarAnalisis();
    }).catch(function (e) { teks("an-info", e.message); })
      .then(function () { b.disabled = false; });
  });

  function gambarAnalisis() {
    var d = dataAn;
    var info = (d.workbook ? "Workbook saat ini" : d.gabungan ? "Gabungan semua grup" : "Kiriman v" + d.versi) +
      " · periode " + (d.tanggal || "—") + " · " + (d.kodeKC || "—");
    if (d.gabungan && d.tanpaData && d.tanpaData.length) info += " · tanpa data (tidak ikut): " + d.tanpaData.join(", ");
    if (d.gabungan) info += " · PD tidak dijumlahkan antar grup, jadi kolom PD resmi tidak ditampilkan.";
    teks("an-info", info);
    tampil("kartu-an-migrasi", !!d.migrasi);
    tampil("kartu-an-netflow", !!d.netflow);
    if (d.migrasi) siapkanMigrasi();
    if (d.netflow) siapkanNetFlow();
  }

  // ---------------- PD Migration ----------------
  function siapkanMigrasi() {
    var tw = dataAn.migrasi.triwulan.filter(function (t) { return t.ada; });
    var w = $("an-tw");
    w.innerHTML = "";
    var pilihan = tw.map(function (t) { return t.nama; });
    if (tw.length > 1) pilihan.push("Gabungan");
    if (!pilihan.length) {
      $("an-mg-sankey").innerHTML = "";
      teks("an-mg-ringkas", "Matriks PD Migration belum terisi.");
      return;
    }
    if (pilihan.indexOf(twAktif) < 0) twAktif = tw[tw.length - 1].nama;   // bawaan: triwulan terakhir yang terisi
    pilihan.forEach(function (n) {
      var b = el("button", "", n.replace("Triwulan ", "TW "));
      b.type = "button";
      b.setAttribute("aria-pressed", n === twAktif ? "true" : "false");
      b.addEventListener("click", function () { twAktif = n; siapkanMigrasi(); });
      w.appendChild(b);
    });
    var dipakai = twAktif === "Gabungan" ? tw : tw.filter(function (t) { return t.nama === twAktif; });
    var mat = [0, 1, 2, 3, 4].map(function () { return [0, 0, 0, 0, 0, 0, 0]; }), awal = [0, 0, 0, 0, 0];
    dipakai.forEach(function (t) {
      for (var i = 0; i < 5; i++) {
        awal[i] += t.saldoAwal[i] || 0;
        for (var j = 0; j < 7; j++) mat[i][j] += (t.matriks[i] || [])[j] || 0;
      }
    });
    gambarMigrasi(mat, awal);
  }

  function kelasMigrasi(i, j) {
    if (j === 6) return ["keluar", "Lunas / keluar"];
    if (j === 5) return ["memburuk", "Hapus buku"];
    return j < i ? ["membaik", "Membaik"] : j === i ? ["tetap", "Tetap"] : ["memburuk", "Memburuk"];
  }

  function gambarMigrasi(mat, awal) {
    var aliran = [], jumlah = { membaik: 0, tetap: 0, memburuk: 0, keluar: 0 }, hb = 0, total = 0;
    var tanpaLancar = $("an-mg-tanpa").checked;
    for (var i = 0; i < 5; i++) {
      total += awal[i];
      for (var j = 0; j < 7; j++) {
        var v = mat[i][j];
        if (!(v > 0)) continue;
        var k = kelasMigrasi(i, j);
        jumlah[k[0]] += v;
        if (j === 5) hb += v;
        if (tanpaLancar && i === 0 && j === 0) continue;   // hanya disembunyikan dari gambar; persentase tetap dari saldo penuh
        aliran.push({ dari: "k" + i, ke: j < 5 ? "t" + j : j === 5 ? "hb" : "lain", nilai: v, kelas: k[0], namaKelas: k[1],
                      info: pct(awal[i] ? v / awal[i] : null) + " dari saldo awal " + LABEL_KOL[i] });
      }
    }
    teks("an-mg-ringkas", "Saldo awal Rp " + rp(total) + " · memburuk " + pct(total ? jumlah.memburuk / total : null) +
      " (termasuk hapus buku " + pct(total ? hb / total : null) + ") · membaik " + pct(total ? jumlah.membaik / total : null) +
      " · lunas/keluar " + pct(total ? jumlah.keluar / total : null) +
      (tanpaLancar ? " · tidak digambar: Kol 1 tetap Kol 1 Rp " + rp(mat[0][0]) : ""));
    Grafik.sankey($("an-mg-sankey"), {
      kiri: LABEL_KOL.map(function (l, i) { return { id: "k" + i, label: l }; }),
      kanan: LABEL_KOL.map(function (l, i) { return { id: "t" + i, label: l }; })
        .concat([{ id: "hb", label: "Hapus buku", kelas: "khusus" }, { id: "lain", label: "Lunas/keluar", kelas: "khusus" }]),
      aliran: aliran, tinggi: 300, lebarLabelKiri: 54, lebarLabelKanan: 86,
      formatNilai: function (v) { return "Rp " + rp(v); },
      // label kiri = saldo awal penuh (bukan hanya yang digambar)
      subLabel: function (id, v, kiri) { return kiri ? miliar(awal[+id.substr(1)]) : ""; },
      label: "Sankey perpindahan saldo antarkualitas " + twAktif
    });

    // tabel matriks (% terhadap saldo awal per baris) + PD resmi
    var pdMg = dataAn.gabungan ? null : dataAn.pdMigrasi;
    var t = el("table", "tabel-kualitas tabel-matriks");
    var h = el("tr");
    ["Awal", "Saldo awal", "K1", "K2", "K3", "K4", "K5", "HB", "Lain"].concat(pdMg ? ["PD"] : [])
      .forEach(function (x) { h.appendChild(el("th", "", x)); });
    t.appendChild(h);
    for (var r = 0; r < 5; r++) {
      var tr = el("tr");
      tr.appendChild(el("td", "", LABEL_KOL[r]));
      tr.appendChild(el("td", "", miliar(awal[r])));
      for (var c = 0; c < 7; c++) tr.appendChild(el("td", c === r ? "diagonal" : "", awal[r] ? pct(mat[r][c] / awal[r]) : "—"));
      if (pdMg) tr.appendChild(el("td", "", pct(pdMg[r])));
      t.appendChild(tr);
    }
    var wt = $("an-mg-tabel");
    wt.innerHTML = "";
    wt.appendChild(t);
    wt.appendChild(el("p", "teks-kecil", "Persentase terhadap saldo awal tiap kualitas. HB = hapus buku (KC2900), Lain = lunas/keluar tanpa hapus buku." +
      (pdMg ? " PD = PD Migration resmi dari sheet B. CKPN - KOL INDV." : "")));
  }

  // ---------------- PD Net Flow ----------------
  function totalBulan(nf, m) {
    var t = 0;
    nf.saldo.forEach(function (b) { t += b[m] || 0; });
    return t + (nf.wo[m] || 0);
  }

  function siapkanNetFlow() {
    var nf = dataAn.netflow, sel = $("an-bulan"), lama = sel.value;
    sel.innerHTML = "";
    var pasangan = [];
    for (var m = 0; m + 1 < nf.bulan.length; m++)
      if (totalBulan(nf, m) > 0 && totalBulan(nf, m + 1) > 0) pasangan.push(m);
    pasangan.forEach(function (m) {
      var o = el("option", "", nf.bulan[m] + " → " + nf.bulan[m + 1]);
      o.value = String(m);
      sel.appendChild(o);
    });
    if (!pasangan.length) {
      $("an-nf-sankey").innerHTML = "";
      $("an-nf-tabel").innerHTML = "";
      teks("an-nf-ringkas", "Saldo bucket Net Flow belum terisi untuk dua bulan berurutan.");
      return;
    }
    sel.value = lama !== "" && pasangan.indexOf(Number(lama)) >= 0 ? lama : String(pasangan[pasangan.length - 1]);
    gambarNetFlow();
  }

  function gambarNetFlow() {
    var nf = dataAn.netflow, m = Number($("an-bulan").value), nol = $("an-nol").checked;
    var B = nf.bucket, aliran = [], baris = [], totalGulir = 0, totalAwal = 0;
    var mulai = nol ? 0 : 1;
    // Asumsi net flow: saldo bucket b+1 bulan berikutnya berasal dari bucket b bulan ini.
    // Bucket terakhir (> 360) digabung dengan hapus buku, sama seperti baris 17 sheet B1.
    for (var b = mulai; b < 13; b++) {
      var asal = nf.saldo[b][m] || 0;
      var tujuan = (nf.saldo[b + 1][m + 1] || 0) + (b + 1 === 13 ? (nf.wo[m + 1] || 0) : 0);
      var rate = asal > 0 ? tujuan / asal : null;
      var gulir = Math.min(tujuan, asal);
      totalAwal += asal; totalGulir += gulir;
      if (gulir > 0) aliran.push({ dari: "b" + b, ke: "n" + (b + 1), nilai: gulir, kelas: "memburuk", namaKelas: "Bergulir",
                                   info: "roll rate " + pct(rate) + (rate > 1 ? " (melebihi 100%, pita dibatasi saldo asal)" : "") });
      if (asal - gulir > 0) aliran.push({ dari: "b" + b, ke: "tidak", nilai: asal - gulir, kelas: "keluar", namaKelas: "Tidak bergulir",
                                          info: pct(asal ? (asal - gulir) / asal : null) + " dari saldo " + B[b] });
      baris.push({ b: b, asal: asal, tujuan: tujuan, rate: rate });
    }
    var kiri = [], kanan = [];
    for (var i = mulai; i < 13; i++) kiri.push({ id: "b" + i, label: B[i] });
    for (var j = mulai + 1; j <= 13; j++) kanan.push({ id: "n" + j, label: j === 13 ? "> 360 + HB" : B[j] });
    kanan.push({ id: "tidak", label: "Tidak bergulir", kelas: "khusus" });

    teks("an-nf-ringkas", nf.bulan[m] + " → " + nf.bulan[m + 1] + " · saldo " + (nol ? "" : "bertunggakan ") + "Rp " + rp(totalAwal) +
      " · bergulir " + pct(totalAwal ? totalGulir / totalAwal : null));
    Grafik.sankey($("an-nf-sankey"), {
      kiri: kiri, kanan: kanan, aliran: aliran, tinggi: 380, lebarLabelKiri: 56, lebarLabelKanan: 90,
      formatNilai: function (v) { return "Rp " + rp(v); },
      label: "Sankey perguliran bucket tunggakan " + nf.bulan[m] + " ke " + nf.bulan[m + 1]
    });

    var pdNf = dataAn.gabungan ? null : dataAn.pdNetFlow;
    var t = el("table", "tabel-kualitas");
    var h = el("tr");
    ["Bucket", nf.bulan[m] + " (jt)", "Berikutnya (jt)", "Roll rate"].concat(pdNf ? ["PD"] : []).forEach(function (x) { h.appendChild(el("th", "", x)); });
    t.appendChild(h);
    baris.forEach(function (r) {
      var tr = el("tr");
      tr.appendChild(el("td", "", B[r.b]));
      tr.appendChild(el("td", "", juta(r.asal)));
      tr.appendChild(el("td", "", juta(r.tujuan)));
      tr.appendChild(el("td", r.rate > 1 ? "kurang" : "", pct(r.rate)));
      if (pdNf) tr.appendChild(el("td", "", pct(pdNf[r.b])));
      t.appendChild(tr);
    });
    if (pdNf) {
      var tr13 = el("tr");
      tr13.appendChild(el("td", "", B[13]));
      tr13.appendChild(el("td", "", juta(nf.saldo[13][m])));
      tr13.appendChild(el("td", "", "—")); tr13.appendChild(el("td", "", "—"));
      tr13.appendChild(el("td", "", pct(pdNf[13])));
      t.appendChild(tr13);
    }
    var wt = $("an-nf-tabel");
    wt.innerHTML = "";
    wt.appendChild(t);
    wt.appendChild(el("p", "teks-kecil", "Roll rate = saldo bucket berikutnya pada " + nf.bulan[m + 1] + " ÷ saldo bucket ini pada " + nf.bulan[m] +
      " (satu pasang bulan). Bucket > 360 + hapus buku bersifat akumulatif, sehingga roll rate langkah terakhir bisa di atas 100%; " +
      "di sankey pitanya dibatasi saldo asal. PD resmi dihitung di sheet dari seluruh bulan, jadi bisa berbeda dari roll rate satu bulan ini."));
  }

  $("an-bulan").addEventListener("change", gambarNetFlow);
  $("an-mg-tanpa").addEventListener("change", function () { if (dataAn && dataAn.migrasi) siapkanMigrasi(); });
  $("an-nol").addEventListener("change", function () { if (dataAn && dataAn.netflow) gambarNetFlow(); });

  // gambar ulang saat lebar panel berubah (task pane bisa digeser)
  var tundaUkur = null;
  window.addEventListener("resize", function () {
    clearTimeout(tundaUkur);
    tundaUkur = setTimeout(function () {
      Grafik.sembunyiTip();
      var analisis = !$("tab-analisis").hidden;
      if (analisis && !$("sub-tren").hidden && dataRw) gambarRiwayat();
      if (analisis && !$("sub-komposisi").hidden && dataAn) { if (dataAn.migrasi) siapkanMigrasi(); if (dataAn.netflow) gambarNetFlow(); }
    }, 200);
  });

  // =====================================================================
  // 10b. Mengapa CKPN berubah? (Tahap 4g)
  //   Membandingkan periode terpilih dengan periode sebelumnya: CKPN (individu/kolektif),
  //   PPKA, dan indikator portofolio dari snapshot Overview (OS, EAD, Kol 2–5, NPF).
  // =====================================================================
  var dataPj = null;
  var fmtJt = new Intl.NumberFormat("id-ID", { maximumFractionDigits: 0 });
  var AMBANG_PERSEN = 0.005;   // perubahan < 0,5% dianggap relatif tetap
  var AMBANG_POIN = 0.0005;    // rasio: < 0,05 poin dianggap tetap

  function jt(v) { return v === null || v === undefined || !isFinite(v) ? "—" : fmtJt.format(v / 1e6) + " jt"; }
  function jtTanda(v) { return v === null || v === undefined || !isFinite(v) ? "—" : (v > 0 ? "+" : v < 0 ? "−" : "") + fmtJt.format(Math.abs(v) / 1e6) + " jt"; }
  function poin(v) { return (v > 0 ? "+" : v < 0 ? "−" : "") + (Math.abs(v) * 100).toFixed(2).replace(".", ",") + " poin"; }

  function muatPenjelasan(tanggal) {
    teks("pj-info", "Memuat…");
    return panggil("penjelasanCKPN", { tanggal: tanggal || "" }).then(function (d) {
      dataPj = d;
      var sel = $("pj-periode"), selG = $("pj-lingkup"), lamaG = selG.value;
      sel.innerHTML = "";
      (d.daftarPeriode || []).slice().reverse().forEach(function (t) {
        var o = el("option", "", labelPeriode(t)); o.value = t; sel.appendChild(o);
      });
      if (d.kini) sel.value = d.kini.tanggal;
      selG.innerHTML = "";
      var b = el("option", "", "Bank (semua grup + ABA)"); b.value = ""; selG.appendChild(b);
      (d.grup || []).forEach(function (g) { var o = el("option", "", g); o.value = g; selG.appendChild(o); });
      selG.value = (d.grup || []).indexOf(lamaG) >= 0 ? lamaG : "";
      gambarPenjelasan();
    }).catch(function (e) { teks("pj-info", e.message); });
  }

  // Angka satu periode untuk lingkup terpilih (bank / satu grup)
  function angkaPeriode(p, grup) {
    if (!p) return null;
    var r = p.riwayat, ov = p.overview, metode = r.metode || "nf", mig = metode === "mig";
    var a = { metode: metode };
    var kc;
    if (!grup) {
      a.ckpn = r.ckpn !== null && r.ckpn !== undefined ? r.ckpn : (mig ? r.mig : r.nf);
      a.ind = mig ? r.migInd : r.nfInd;
      a.kol = mig ? r.migKol : r.nfKol;
      a.aba = r.aba ? r.aba.ckpn : 0;
      a.ppka = r.ppka;
      kc = ["KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100"];
    } else {
      var g = null;
      r.grup.forEach(function (x) { if (x.nama === grup && x.ada) g = x; });
      if (!g) return { kosong: true };
      a.ckpn = mig ? g.mig : g.nf;
      a.ind = mig ? g.migInd : g.nfInd;
      a.kol = mig ? g.migKol : g.nfKol;
      a.ppka = g.ppka;
      a.lgd = g.lgd;
      kc = g.kodeKC.split(",");
    }
    a.kc = kc;
    if (ov && ov.segmen) {
      var os = 0, ead = 0, k25 = 0, npf = 0;
      ov.segmen.forEach(function (s) {
        if (kc.indexOf(s.kode) < 0) return;
        os += s.totalOS || 0; ead += s.totalEAD || 0;
        k25 += (s.os[1] || 0) + (s.os[2] || 0) + (s.os[3] || 0) + (s.os[4] || 0);
        npf += (s.os[2] || 0) + (s.os[3] || 0) + (s.os[4] || 0);
      });
      a.os = os; a.ead = ead; a.k25 = k25; a.npf = npf;
      a.npfPct = os ? npf / os : null; a.k25Pct = os ? k25 / os : null;
    }
    return a;
  }

  function arahNilai(lalu, kini) {
    if (lalu === null || lalu === undefined || kini === null || kini === undefined) return 0;
    var d = kini - lalu, basis = Math.abs(lalu) || Math.abs(kini) || 1;
    return Math.abs(d) / basis < AMBANG_PERSEN ? 0 : (d > 0 ? 1 : -1);
  }
  function arahRasio(lalu, kini) {
    if (lalu === null || lalu === undefined || kini === null || kini === undefined) return 0;
    var d = kini - lalu;
    return Math.abs(d) < AMBANG_POIN ? 0 : (d > 0 ? 1 : -1);
  }
  var PANAH = { "1": "▲ naik", "-1": "▼ turun", "0": "≈ tetap" };

  function gambarPenjelasan() {
    var d = dataPj, w = $("pj-isi");
    w.innerHTML = "";
    $("pj-ov").innerHTML = "";
    if (!d || !d.kini) { teks("pj-info", "Belum ada periode dengan kiriman grup tersimpan."); return; }
    var grup = $("pj-lingkup").value;
    if (!d.lalu) {
      teks("pj-info", labelPeriode(d.kini.tanggal) + " adalah periode pertama di database — belum ada pembanding.");
      return;
    }
    teks("pj-info", labelPeriode(d.lalu.tanggal) + " → " + labelPeriode(d.kini.tanggal) + " · " + (grup || "bank") +
      " · CKPN memakai metode tahunan periode masing-masing");

    // ---- snapshot overview yang belum ada ----
    var kurang = [d.lalu, d.kini].filter(function (p) { return !p.overview; });
    if (kurang.length) {
      var box = el("div", "peringatan-box");
      box.appendChild(el("div", "", "Snapshot Overview belum ada untuk " + kurang.map(function (p) { return labelPeriode(p.tanggal); }).join(" dan ") +
        ", sehingga OS, EAD, dan NPF belum bisa dibandingkan. Periode Master bisa diisi dengan membuka Ringkasan › Muat data; " +
        "periode lain dari file template-nya."));
      if (d.bolehMenulis) kurang.forEach(function (p) {
        box.appendChild(tombol("Ambil overview " + labelPeriode(p.tanggal) + " dari file template…", "tautan", function () {
          panggil("overviewDariFile", { tanggal: p.tanggal }, 600000).then(function (r) {
            if (!r.batal) muatPenjelasan($("pj-periode").value);
          }).catch(function (e) { alert(e.message); });
        }));
      });
      $("pj-ov").appendChild(box);
    }

    var A = angkaPeriode(d.lalu, grup), B = angkaPeriode(d.kini, grup);
    if (!A || !B || A.kosong || B.kosong) {
      w.appendChild(el("div", "teks-kecil", "Grup " + grup + " tidak tersimpan di salah satu periode."));
      return;
    }

    // ---- tabel indikator ----
    var baris = [
      { label: "CKPN (" + (B.metode === "mig" ? "Migration" : "Net Flow") + ")", a: A.ckpn, b: B.ckpn, utama: true },
      { label: "· CKPN Individu", a: A.ind, b: B.ind },
      { label: "· CKPN Kolektif", a: A.kol, b: B.kol }
    ];
    if (!grup) baris.push({ label: "· CKPN ABA", a: A.aba, b: B.aba });
    baris.push({ label: "PPKA OJK", a: A.ppka, b: B.ppka });
    if (A.os !== undefined && B.os !== undefined) {
      baris.push({ label: "OS pembiayaan", a: A.os, b: B.os, driver: "os" });
      baris.push({ label: "EAD", a: A.ead, b: B.ead, driver: "ead" });
      baris.push({ label: "OS Kol 2–5", a: A.k25, b: B.k25, driver: "k25" });
      baris.push({ label: "OS NPF (Kol 3–5)", a: A.npf, b: B.npf, driver: "npf" });
      baris.push({ label: "Kol 2–5 %", a: A.k25Pct, b: B.k25Pct, rasio: true, driver: "k25Pct" });
      baris.push({ label: "NPF %", a: A.npfPct, b: B.npfPct, rasio: true, driver: "npfPct" });
    }
    if (grup) baris.push({ label: "LGD gabungan", a: A.lgd, b: B.lgd, rasio: true, driver: "lgd" });

    var t = el("table", "tabel-kualitas tabel-pj");
    var h = el("tr");
    ["Indikator", labelPeriode(d.lalu.tanggal), labelPeriode(d.kini.tanggal), "Perubahan", ""].forEach(function (x) { h.appendChild(el("th", "", x)); });
    t.appendChild(h);
    baris.forEach(function (r) {
      r.arah = r.rasio ? arahRasio(r.a, r.b) : arahNilai(r.a, r.b);
      var tr = el("tr", r.utama ? "total" : "");
      tr.appendChild(el("td", "", r.label));
      tr.appendChild(el("td", "", r.rasio ? pct(r.a) : jt(r.a)));
      tr.appendChild(el("td", "", r.rasio ? pct(r.b) : jt(r.b)));
      var delta = r.a === null || r.b === null || r.a === undefined || r.b === undefined ? null : r.b - r.a;
      tr.appendChild(el("td", "", delta === null ? "—" : r.rasio ? poin(delta) :
        jtTanda(delta) + (r.a ? " (" + (delta >= 0 ? "+" : "−") + pct(Math.abs(delta / r.a)) + ")" : "")));
      tr.appendChild(el("td", "arah arah" + r.arah, PANAH[String(r.arah)]));
      t.appendChild(tr);
    });

    // ---- narasi ----
    var nar = el("div", "narasi-pj");
    var dC = B.ckpn - A.ckpn, arahC = arahNilai(A.ckpn, B.ckpn);
    nar.appendChild(el("p", "narasi-judul", "CKPN " + (grup || "bank") + " " + (arahC > 0 ? "naik" : arahC < 0 ? "turun" : "relatif tetap") +
      " " + jtTanda(dC) + (A.ckpn ? " (" + (dC >= 0 ? "+" : "−") + pct(Math.abs(dC / A.ckpn)) + ")" : "") +
      " dari " + labelPeriode(d.lalu.tanggal) + " ke " + labelPeriode(d.kini.tanggal) + "."));
    var ul = el("ul", "narasi-daftar");
    function poinNarasi(s) { ul.appendChild(el("li", "", s)); }

    // sumber perubahan: individu vs kolektif (vs ABA)
    var dI = (B.ind || 0) - (A.ind || 0), dK = (B.kol || 0) - (A.kol || 0), dA = grup ? 0 : (B.aba || 0) - (A.aba || 0);
    var komponen = [["Individu", dI], ["Kolektif", dK]].concat(grup ? [] : [["ABA", dA]]);
    komponen.sort(function (x, y) { return Math.abs(y[1]) - Math.abs(x[1]); });
    if (arahC !== 0)
      poinNarasi("Sumber perubahan terbesar: CKPN " + komponen[0][0] + " " + jtTanda(komponen[0][1]) +
        komponen.slice(1).map(function (k) { return "; " + k[0] + " " + jtTanda(k[1]); }).join("") + ".");

    // indikator portofolio yang searah / berlawanan dengan arah CKPN
    var nama = { os: "OS", ead: "EAD", k25: "OS Kol 2–5", npf: "OS NPF", k25Pct: "rasio Kol 2–5", npfPct: "NPF", lgd: "LGD" };
    var searah = [], lawan = [];
    baris.forEach(function (r) {
      if (!r.driver || r.arah === 0) return;
      var teksR = nama[r.driver] + " " + (r.arah > 0 ? "naik " : "turun ") +
        (r.rasio ? poin(r.b - r.a).replace(/^[+−]/, "") + " (" + pct(r.a) + " → " + pct(r.b) + ")" : jtTanda(r.b - r.a).replace(/^[+−]/, "") +
         (r.a ? " (" + pct(Math.abs((r.b - r.a) / r.a)) + ")" : ""));
      if (arahC !== 0 && r.arah === arahC) searah.push(teksR); else lawan.push(teksR);
    });
    if (arahC !== 0 && searah.length) poinNarasi("Searah dengan " + (arahC > 0 ? "kenaikan" : "penurunan") + " CKPN: " + searah.join("; ") + ".");
    if (lawan.length) poinNarasi((arahC !== 0 ? "Berlawanan / meredam: " : "Pergerakan indikator: ") + lawan.join("; ") + ".");
    if (A.os === undefined || B.os === undefined) poinNarasi("Indikator OS/EAD/NPF belum tersedia — lengkapi snapshot Overview.");

    // petunjuk sebab
    var npfR = baris.filter(function (r) { return r.driver === "npfPct" || r.driver === "k25Pct" || r.driver === "npf" || r.driver === "k25"; });
    var kualitasSearah = npfR.some(function (r) { return r.arah === arahC && arahC !== 0; });
    if (arahC !== 0 && Math.abs(dK) >= Math.abs(dI) && !kualitasSearah && npfR.length)
      poinNarasi("Perubahan kolektif tidak diikuti pergerakan kualitas (NPF / Kol 2–5) yang searah — kemungkinan berasal dari perubahan PD atau LGD. " +
        "Cek sankey PD di bawah" + (grup ? " dan baris LGD." : "."));
    if (arahC !== 0 && Math.abs(dI) > Math.abs(dK))
      poinNarasi("Perubahan terutama dari CKPN Individu — lihat rincian kontrak di bawah (masuk/keluar Top-N atau perubahan agunan).");
    nar.appendChild(ul);
    w.appendChild(nar);
    var gul = el("div", "gulir-x"); gul.appendChild(t); w.appendChild(gul);
    w.appendChild(el("p", "teks-kecil", "Angka dalam juta Rp. ▲/▼ bila berubah ≥ 0,5% (rasio: ≥ 0,05 poin). " +
      "OS, EAD, Kol 2–5, dan NPF dari snapshot Overview (template OJK), dijumlah untuk KC dalam lingkup."));

    // ---- kontrak individu ----
    var ind = (d.individu && d.individu.kontrak) || [];
    var saring = grup ? A.kc.concat(B.kc) : null;
    var adaKodeKc = ind.some(function (k) { return /^KC\d{4}$/.test(k.kc); });
    var daftar = ind.filter(function (k) { return !saring || !adaKodeKc || saring.indexOf(k.kc) >= 0; })
      .map(function (k) { k.delta = k.kini - k.lalu; return k; })
      .filter(function (k) { return Math.abs(k.delta) >= 1; });
    if (daftar.length) {
      var masuk = daftar.filter(function (k) { return k.jenis === "masuk"; }), keluar = daftar.filter(function (k) { return k.jenis === "keluar"; });
      var jml = function (l) { return l.reduce(function (s, k) { return s + k.delta; }, 0); };
      var det = el("details", "jarak-atas");
      det.appendChild(el("summary", "teks-kecil tautan-ringkas", "Rincian CKPN Individu: " + masuk.length + " kontrak masuk (" + jtTanda(jml(masuk)) + "), " +
        keluar.length + " keluar (" + jtTanda(jml(keluar)) + "), " + (daftar.length - masuk.length - keluar.length) + " berubah"));
      daftar.sort(function (x, y) { return Math.abs(y.delta) - Math.abs(x.delta); });
      var ti = el("table", "tabel-kualitas tabel-pj");
      var hi = el("tr");
      ["Kontrak", "Status", "CKPN lalu", "CKPN kini", "Δ"].forEach(function (x) { hi.appendChild(el("th", "", x)); });
      ti.appendChild(hi);
      daftar.slice(0, 15).forEach(function (k) {
        var tr = el("tr");
        var c = el("td");
        c.appendChild(el("div", "", k.nama || k.kontrak));
        c.appendChild(el("div", "teks-kecil", k.kontrak + (k.kc ? " · " + k.kc : "")));
        tr.appendChild(c);
        tr.appendChild(el("td", "", k.jenis === "tetap" ? "berubah" : k.jenis));
        tr.appendChild(el("td", "", jt(k.lalu)));
        tr.appendChild(el("td", "", jt(k.kini)));
        tr.appendChild(el("td", "", jtTanda(k.delta)));
        ti.appendChild(tr);
      });
      var gul2 = el("div", "gulir-x"); gul2.appendChild(ti); det.appendChild(gul2);
      if (daftar.length > 15) det.appendChild(el("p", "teks-kecil", "15 perubahan terbesar dari " + daftar.length + "."));
      w.appendChild(det);
    }
  }

  $("pj-periode").addEventListener("change", function () { muatPenjelasan(this.value); });
  $("pj-lingkup").addEventListener("change", gambarPenjelasan);

  // =====================================================================
  // Mode PD & LGD setahun sekali (Tahap 5e) — acuan Desember tahun lalu
  // =====================================================================
  var infoAcuan = null;

  function muatAcuan() {
    return panggil("infoAcuan").then(function (d) { infoAcuan = d; gambarAcuan(); gambarPengaturan(); })
      .catch(function () { infoAcuan = null; gambarAcuan(); gambarPengaturan(); });
  }

  function acuanGrup(kode) {
    var g = null;
    ((infoAcuan && infoAcuan.grup) || []).forEach(function (x) { if (x.kodeKC === kode) g = x; });
    return g;
  }

  function pakaiAcuan() {
    if (!infoAcuan || !infoAcuan.tahunan || infoAcuan.desember || !infoAcuan.tahunAcuan) return false;
    var r = document.querySelector('input[name="acuan-mode"]:checked');
    return !!r && r.value === "acuan";
  }

  function teksAcuanGrup(kode) {
    var g = acuanGrup(kode);
    if (g && g.acuan) return "PD & LGD: acuan " + g.acuan.label + " (status " + g.acuan.status + ") — hanya CKPN Individu & Summary yang dihitung.";
    return "PD & LGD: acuan Desember " + infoAcuan.tahunAcuan + " belum ada untuk grup ini — semua langkah dihitung penuh.";
  }

  function ringkasAcuan() {
    var ada = 0, tidak = [];
    ((infoAcuan && infoAcuan.grup) || []).forEach(function (g) { if (g.acuan) ada++; else tidak.push(g.nama); });
    return ada + " grup memakai acuan Desember " + infoAcuan.tahunAcuan +
      (tidak.length ? "; dihitung penuh (acuan belum ada): " + tidak.join(", ") : "") + ".";
  }

  function gambarAcuan() {
    var d = infoAcuan, kotak = $("pg-acuan");
    if (!d || !d.tahunan) { kotak.hidden = true; return; }
    kotak.hidden = false;
    var ul = $("acuan-grup");
    ul.innerHTML = "";
    if (d.desember || !d.tahunAcuan) {
      teks("acuan-judul", "Metode PD & LGD: setahun sekali · Desember");
      tampil("acuan-pilihan", false);
      ul.appendChild(el("li", "", "Laporan Desember: PD & LGD dihitung penuh. Hasilnya menjadi acuan PD & LGD untuk Januari–November tahun berikutnya."));
      return;
    }
    teks("acuan-judul", "Metode PD & LGD: setahun sekali (Master!D17) · acuan Desember " + d.tahunAcuan);
    tampil("acuan-pilihan", true);
    (d.grup || []).forEach(function (g) {
      var li = el("li", g.acuan ? "" : "kurang");
      li.appendChild(el("b", "", g.nama + ": "));
      li.appendChild(el("span", "", g.acuan
        ? "acuan " + g.acuan.label + " · " + g.acuan.status + " · LGD " + (g.acuan.lgd * 100).toFixed(2).replace(".", ",") + "%"
        : (g.masalah || "acuan belum ada") + " → dihitung penuh"));
      ul.appendChild(li);
    });
  }

  // =====================================================================
  // Lihat & koreksi data tersimpan satu versi grup (Tahap 5b)
  //   Individu : CKPN = PN "Ya" ? MAX(0, OS − (agunan − biaya)) : 0   (kolom K sheet)
  //   LGD CS   : LGD = 1 − ΣG/ΣC; gabungan = 1 − (ΣG + rec ER)/(ΣC + WO ER)
  //   Kolektif : × LGD gabungan baru / lama (sheet KOL INDV memakai satu sel LGD dari B4)
  // Server (KoreksiRun.cs) menghitung ulang dengan rumus yang sama saat menyimpan.
  // =====================================================================
  var kr = null;

  function bukaKoreksi(runId, nama) {
    // editor ada di Hitung › Per grup; dibuka juga dari Analisis › Versi dan kartu Hitung semua grup
    if (tabAktif !== "hitung" || subAktif.hitung !== "pergrup") pindahTab("hitung", "pergrup", true);
    tampil("kartu-koreksi", true);
    teks("kr-judul", "Data " + (nama || ""));
    teks("kr-sub", "Memuat…");
    $("kr-hasil").innerHTML = "";
    return panggil("koreksiData", { runId: runId }).then(function (d) {
      kr = { d: d, nama: nama, tab: (kr && kr.tab) || "individu", ind: {}, cs: {}, hapus: {}, tambah: [] };
      d.individu.forEach(function (b) { kr.ind[b.id] = { jaminan: b.jaminan, biaya: b.biaya }; });
      d.lgdcs.forEach(function (c) {
        var h = d.lgd.haircut;
        var rumus = h !== null && h !== undefined && c.agunan > 0 &&
          Math.abs(c.recovery - Math.min(c.agunan * (1 - h), c.pokok)) < 1;
        kr.cs[c.id] = { agunan: c.agunan, recovery: c.recovery, rumus: rumus, manual: false };
      });
      $("kr-alasan").value = "";
      gambarKoreksi();
      $("kartu-koreksi").scrollIntoView({ block: "start" });
    }).catch(function (e) { teks("kr-sub", e.message); });
  }

  function bisaEdit() { return kr && kr.d.bolehKoreksi; }
  function bisaEditCs() { return bisaEdit() && !kr.d.lgd.masalah; }

  function ckpnIndividu(b, nilai) {
    return String(b.pn || "").trim().toLowerCase() === "ya" ? Math.max(0, b.os - (nilai.jaminan - nilai.biaya)) : 0;
  }
  function beda(a, b) { return Math.abs((a || 0) - (b || 0)) >= 0.005; }

  function hitungKoreksi() {
    var d = kr.d, rg = d.ringkasan, lgd = d.lgd;
    var indLama = 0, indBaru = 0, nInd = 0;
    d.individu.forEach(function (b) {
      var n = kr.ind[b.id], ubah = beda(n.jaminan, b.jaminan) || beda(n.biaya, b.biaya);
      indLama += b.ckpn;
      indBaru += ubah ? ckpnIndividu(b, n) : b.ckpn;
      if (ubah) nInd++;
    });
    var wo = 0, rec = 0, nCs = 0;
    d.lgdcs.forEach(function (c) {
      var n = kr.cs[c.id];
      if (kr.hapus[c.id]) { nCs++; return; }
      if (beda(n.agunan, c.agunan) || beda(n.recovery, c.recovery)) nCs++;
      wo += c.pokok; rec += n.recovery;
    });
    kr.tambah.forEach(function (t) { nCs++; wo += t.pokok; rec += t.recovery; });

    function L(w, r) { return w === 0 ? 0 : 1 - r / w; }
    var gabLama = L(lgd.woCs + lgd.woEr, lgd.recCs + lgd.recEr);
    var gabBaru = nCs ? L(wo + lgd.woEr, rec + lgd.recEr) : gabLama;
    var faktor = nCs && gabLama > 0 ? gabBaru / gabLama : 1;
    var dInd = indBaru - indLama;
    function ada(k) { return rg[k] !== undefined && rg[k] !== null; }
    var h = {
      nInd: nInd, nCs: nCs, faktor: faktor,
      baris: [
        { label: "CKPN Individu", lama: rg.nf_individu, baru: ada("nf_individu") ? rg.nf_individu + dInd : null },
        { label: "LGD CS", lama: rg.lgd_cs, baru: nCs ? L(wo, rec) : rg.lgd_cs, persen: true },
        { label: "LGD gabungan", lama: rg.lgd_gabungan, baru: nCs ? gabBaru : rg.lgd_gabungan, persen: true },
        { label: "Kolektif Net Flow", lama: rg.nf_kolektif, baru: ada("nf_kolektif") ? rg.nf_kolektif * faktor : null, metode: "nf" },
        { label: "Kolektif Migration", lama: rg.mig_kolektif, baru: ada("mig_kolektif") ? rg.mig_kolektif * faktor : null, metode: "mig" },
        { label: "Total Net Flow", lama: rg.nf_total, metode: "nf", tebal: true,
          baru: ada("nf_individu") && ada("nf_kolektif") ? rg.nf_individu + dInd + rg.nf_kolektif * faktor : rg.nf_total },
        { label: "Total Migration", lama: rg.mig_total, metode: "mig", tebal: true,
          baru: ada("mig_individu") && ada("mig_kolektif") ? rg.mig_individu + dInd + rg.mig_kolektif * faktor : rg.mig_total }
      ]
    };
    return h;
  }

  function gambarKoreksi() {
    var d = kr.d;
    teks("kr-judul", "Data " + (kr.nama || d.kodeKC) + " · v" + d.versi + (d.aktif ? "" : " (tidak aktif)"));
    teks("kr-sub", tanggalPanjang(d.tanggal) + " · " + d.kodeKC + " · disimpan " + d.waktu + " oleh " + d.pengguna +
      (d.catatan ? " · " + d.catatan : ""));
    var tolak = $("kr-tolak");
    tolak.hidden = d.bolehKoreksi;
    tolak.textContent = d.bolehKoreksi ? "" : "Hanya lihat: " + d.alasanTolak;
    document.querySelectorAll("[data-kr]").forEach(function (b) {
      var j = b.getAttribute("data-kr");
      b.setAttribute("aria-pressed", j === kr.tab ? "true" : "false");
      b.textContent = (j === "individu" ? "CKPN Individu (" + d.individu.length + ")" : "LGD CS (" + (d.lgdcs.length + kr.tambah.length) + ")");
    });
    tampil("kr-form", bisaEdit());
    gambarTabelKoreksi();
    gambarDampak();
  }

  function isianAngka(nilai, label, fn) {
    var inp = el("input", "isian angka kr-isian");
    inp.type = "text"; inp.inputMode = "numeric";
    inp.setAttribute("aria-label", label);
    inp.value = fmt.format(nilai || 0);
    inp.addEventListener("input", function () { var v = parseAngka(inp.value); fn(v === null ? 0 : v, false); });
    inp.addEventListener("blur", function () { var v = parseAngka(inp.value); inp.value = fmt.format(v === null ? 0 : v); fn(v === null ? 0 : v, true); });
    return inp;
  }

  function sel(label, isi, cls) {
    var c = el("div", "kr-sel" + (cls ? " " + cls : ""));
    c.appendChild(el("div", "kr-sel-label", label));
    if (typeof isi === "string") c.appendChild(el("div", "kr-sel-nilai", isi)); else c.appendChild(isi);
    return c;
  }

  function gambarTabelKoreksi() {
    var w = $("kr-tabel"), d = kr.d;
    w.innerHTML = "";
    var cari = $("kr-cari").value.trim().toLowerCase();
    var edit = bisaEdit();
    tampil("kr-tambah-wadah", kr.tab === "lgdcs" && bisaEditCs());

    if (kr.tab === "individu") {
      var list = d.individu.filter(function (b) { return !cari || (b.kontrak + " " + b.nama + " " + b.cif + " " + b.kc).toLowerCase().indexOf(cari) >= 0; });
      teks("kr-info-tabel", d.individu.length === 0 ? "Versi ini tidak menyimpan tabel Individu." :
        list.length + " kontrak · CKPN = MAX(0, OS − (agunan − biaya)) bila ada penurunan nilai.");
      list.forEach(function (b) {
        var n = kr.ind[b.id];
        var row = el("div", "kr-baris");
        var ck = el("div", "kr-sel-nilai");
        function segar() {
          var ubah = beda(n.jaminan, b.jaminan) || beda(n.biaya, b.biaya);
          row.classList.toggle("diubah", ubah);
          ck.textContent = ubah ? rp(b.ckpn) + " → " + rp(ckpnIndividu(b, n)) : rp(b.ckpn);
        }
        var kep = el("div", "kr-kepala");
        kep.appendChild(el("b", "", b.kontrak));
        kep.appendChild(el("span", "", " · " + b.nama));
        var sub = el("div", "teks-kecil", b.kc + " · CIF " + b.cif + " · penurunan nilai: " + (b.pn || "—") + (b.disesuaikan ? " · berpenyesuaian" : ""));
        kep.appendChild(sub);
        row.appendChild(kep);
        var g = el("div", "kr-grid");
        g.appendChild(sel("OS", rp(b.os)));
        g.appendChild(sel("Agunan (I)", edit ? isianAngka(n.jaminan, "Agunan " + b.kontrak, function (v) { n.jaminan = v; segar(); gambarDampak(); }) : rp(n.jaminan)));
        g.appendChild(sel("Biaya jual (J)", edit ? isianAngka(n.biaya, "Biaya penjualan " + b.kontrak, function (v) { n.biaya = v; segar(); gambarDampak(); }) : rp(n.biaya)));
        g.appendChild(sel("CKPN (K)", ck, "kr-hasil-sel"));
        row.appendChild(g);
        segar();
        w.appendChild(row);
      });
      return;
    }

    var lgd = d.lgd;
    var info = d.lgdcs.length + kr.tambah.length + " rekening · LGD CS = 1 − Σ realisasi / Σ pokok awal.";
    if (lgd.masalah) info += " Koreksi LGD CS tidak tersedia: " + lgd.masalah;
    else if (lgd.haircut !== null && lgd.haircut !== undefined)
      info += " Realisasi yang masih rumus agunan ikut berubah saat agunan diubah (haircut " + (lgd.haircut * 100).toFixed(1).replace(".", ",") + "%).";
    teks("kr-info-tabel", info);
    var editCs = bisaEditCs();

    d.lgdcs.forEach(function (c) {
      if (cari && (c.rek + " " + c.nama).toLowerCase().indexOf(cari) < 0) return;
      var n = kr.cs[c.id];
      var row = el("div", "kr-baris");
      var sf = el("div", "kr-sel-nilai");
      var inpRec = null;
      function segar() {
        var hapus = !!kr.hapus[c.id];
        row.classList.toggle("dihapus", hapus);
        row.classList.toggle("diubah", !hapus && (beda(n.agunan, c.agunan) || beda(n.recovery, c.recovery)));
        var s = n.recovery > c.pokok ? 0 : c.pokok - n.recovery;
        sf.textContent = beda(s, c.shortfall) ? rp(c.shortfall) + " → " + rp(s) : rp(s);
      }
      var kep = el("div", "kr-kepala baris-antara");
      var kiri = el("div");
      kiri.appendChild(el("b", "", c.rek));
      kiri.appendChild(el("span", "", " · " + (c.nama || "—")));
      kiri.appendChild(el("div", "teks-kecil", "Diserahkan " + (c.thnSerah || "—") + " · eksekusi " + (c.thnEks || "—") +
        (c.sumber && c.sumber !== "sistem" ? " · " + c.sumber : "") + (n.rumus ? " · G rumus agunan" : "")));
      kep.appendChild(kiri);
      if (editCs) {
        var bh = tombol(kr.hapus[c.id] ? "Batalkan hapus" : "Hapus", "tautan hapus", function () {
          if (kr.hapus[c.id]) delete kr.hapus[c.id]; else kr.hapus[c.id] = true;
          bh.textContent = kr.hapus[c.id] ? "Batalkan hapus" : "Hapus";
          segar(); gambarDampak();
        });
        kep.appendChild(bh);
      }
      row.appendChild(kep);
      var g = el("div", "kr-grid");
      g.appendChild(sel("Pokok (C)", rp(c.pokok)));
      g.appendChild(sel("Agunan (D)", editCs ? isianAngka(n.agunan, "Agunan " + c.rek, function (v) {
        n.agunan = v;
        if (n.rumus && !n.manual) {
          n.recovery = Math.min(v * (1 - d.lgd.haircut), c.pokok);
          if (inpRec && document.activeElement !== inpRec) inpRec.value = fmt.format(n.recovery);
        }
        segar(); gambarDampak();
      }) : rp(n.agunan)));
      if (editCs) inpRec = isianAngka(n.recovery, "Realisasi " + c.rek, function (v) { n.recovery = v; n.manual = true; segar(); gambarDampak(); });
      g.appendChild(sel("Realisasi (G)", editCs ? inpRec : rp(n.recovery)));
      g.appendChild(sel("Shortfall (H)", sf, "kr-hasil-sel"));
      row.appendChild(g);
      segar();
      w.appendChild(row);
    });

    kr.tambah.forEach(function (t, i) {
      var row = el("div", "kr-baris baru");
      var kep = el("div", "kr-kepala baris-antara");
      var kiri = el("div");
      kiri.appendChild(el("b", "", t.rek));
      kiri.appendChild(el("span", "", " · " + (t.nama || "—") + " · baris baru"));
      kep.appendChild(kiri);
      kep.appendChild(tombol("Buang", "tautan hapus", function () { kr.tambah.splice(i, 1); gambarKoreksi(); }));
      row.appendChild(kep);
      var g = el("div", "kr-grid");
      g.appendChild(sel("Pokok (C)", rp(t.pokok)));
      g.appendChild(sel("Agunan (D)", rp(t.agunan)));
      g.appendChild(sel("Realisasi (G)", rp(t.recovery)));
      g.appendChild(sel("Shortfall (H)", rp(t.recovery > t.pokok ? 0 : t.pokok - t.recovery), "kr-hasil-sel"));
      row.appendChild(g);
      w.appendChild(row);
    });
  }

  function gambarDampak() {
    var h = hitungKoreksi(), w = $("kr-dampak");
    var metode = statusP && statusP.metode;
    w.innerHTML = "";
    var t = el("table", "tabel-staging tabel-dampak");
    var thead = el("tr");
    ["", "Sebelum", "Sesudah", "Selisih"].forEach(function (x) { thead.appendChild(el("th", x ? "angka" : "", x)); });
    t.appendChild(thead);
    h.baris.forEach(function (b) {
      var tr = el("tr", (b.tebal ? "tebal" : "") + (metode && b.metode === metode ? " metode" : ""));
      tr.appendChild(el("td", "", b.label + (metode && b.metode === metode ? " ★" : "")));
      var fmtV = b.persen ? function (v) { return v === null || v === undefined ? "—" : (v * 100).toFixed(2).replace(".", ",") + "%"; } : rp;
      tr.appendChild(el("td", "angka", fmtV(b.lama)));
      tr.appendChild(el("td", "angka", fmtV(b.baru)));
      var selisih = b.lama === null || b.lama === undefined || b.baru === null || b.baru === undefined ? null : b.baru - b.lama;
      var td = el("td", "angka" + (selisih && Math.abs(selisih) > (b.persen ? 1e-7 : 0.5) ? (selisih > 0 ? " naik" : " turun") : ""),
        selisih === null ? "—" : b.persen ? (Math.abs(selisih) < 1e-7 ? "0" : (selisih > 0 ? "+" : "−") + (Math.abs(selisih) * 100).toFixed(2).replace(".", ",") + " poin")
          : bertanda(Math.round(selisih)));
      tr.appendChild(td);
      t.appendChild(tr);
    });
    w.appendChild(t);

    var cat = ["Perubahan: " + h.nInd + " kontrak Individu, " + h.nCs + " baris LGD CS."];
    if (h.nCs) cat.push("CKPN Kolektif ikut berubah ×" + h.faktor.toFixed(4).replace(".", ",") +
      " karena sheet KOL INDV memakai satu LGD (LGD gabungan B4) untuk Net Flow dan Migration.");
    if (metode) cat.push("★ = metode konsolidasi tahun ini.");
    cat.push("Hanya versi periode ini yang berubah — penyesuaian untuk perhitungan bulan berikutnya tidak ikut diubah.");
    if (statusP && statusP.tanggal === kr.d.tanggal && statusP.kodeKCMaster === kr.d.kodeKC && statusP.periodeMaster)
      cat.push("Sheet di workbook masih berisi angka sebelum koreksi; Simpan grup dari sheet akan membuat versi baru tanpa koreksi ini.");
    teks("kr-catatan", cat.join(" "));
    $("kr-simpan").disabled = !bisaEdit() || (h.nInd + h.nCs) === 0;
    $("kr-simpan").textContent = (h.nInd + h.nCs) ? "Simpan sebagai v" + (kr.d.versi + 1) : "Simpan sebagai versi baru";
  }

  document.querySelectorAll("[data-kr]").forEach(function (b) {
    b.addEventListener("click", function () { kr.tab = b.getAttribute("data-kr"); $("kr-cari").value = ""; gambarKoreksi(); });
  });
  $("kr-cari").addEventListener("input", gambarTabelKoreksi);
  $("kr-tutup").addEventListener("click", function () { tampil("kartu-koreksi", false); kr = null; });
  $("kr-reset").addEventListener("click", function () { if (kr) bukaKoreksi(kr.d.runId, kr.nama); });
  $("kr-tambah-buka").addEventListener("click", function () { tampil("kr-tambah", true); tampil("kr-tambah-buka", false); $("kr-t-rek").focus(); });
  $("kr-t-batal").addEventListener("click", function () { tampil("kr-tambah", false); tampil("kr-tambah-buka", true); });
  $("kr-t-ok").addEventListener("click", function () {
    var rek = $("kr-t-rek").value.trim(), pokok = parseAngka($("kr-t-pokok").value) || 0;
    var ada = kr.d.lgdcs.some(function (c) { return c.rek.toLowerCase() === rek.toLowerCase() && !kr.hapus[c.id]; }) ||
              kr.tambah.some(function (t) { return t.rek.toLowerCase() === rek.toLowerCase(); });
    if (!rek) { teks("kr-t-pesan", "No. rekening wajib diisi."); return; }
    if (ada) { teks("kr-t-pesan", "Rekening " + rek + " sudah ada di tabel."); return; }
    if (pokok <= 0) { teks("kr-t-pesan", "Pokok awal wajib diisi."); return; }
    var thn = $("kr-t-thn").value.split("/");
    kr.tambah.push({ rek: rek, nama: $("kr-t-nama").value.trim(), pokok: pokok,
      agunan: parseAngka($("kr-t-agunan").value) || 0, recovery: parseAngka($("kr-t-rec").value) || 0,
      thnSerah: (thn[0] || "").trim(), thnEks: (thn[1] || "").trim() });
    ["kr-t-rek", "kr-t-nama", "kr-t-pokok", "kr-t-agunan", "kr-t-rec", "kr-t-thn"].forEach(function (id) { $(id).value = ""; });
    $("kr-t-pesan").textContent = "";
    tampil("kr-tambah", false); tampil("kr-tambah-buka", true);
    gambarKoreksi();
  });

  $("kr-simpan").addEventListener("click", function () {
    var d = kr.d, alasan = $("kr-alasan").value.trim();
    if (!alasan) { $("kr-hasil").innerHTML = ""; $("kr-hasil").appendChild(el("div", "peringatan-box", "Alasan koreksi wajib diisi.")); $("kr-alasan").focus(); return; }
    var ind = [], ubah = [], hapus = [];
    d.individu.forEach(function (b) {
      var n = kr.ind[b.id];
      if (beda(n.jaminan, b.jaminan) || beda(n.biaya, b.biaya)) ind.push({ id: b.id, jaminan: n.jaminan, biaya: n.biaya });
    });
    d.lgdcs.forEach(function (c) {
      var n = kr.cs[c.id];
      if (kr.hapus[c.id]) hapus.push(c.id);
      else if (beda(n.agunan, c.agunan) || beda(n.recovery, c.recovery)) ubah.push({ id: c.id, agunan: n.agunan, recovery: n.recovery });
    });
    var b = this;
    b.disabled = true;
    panggil("koreksiSimpan", { runId: d.runId, alasan: alasan, individu: ind,
      lgdcs: { ubah: ubah, hapus: hapus, tambah: kr.tambah } }, 120000).then(function (r) {
      var nama = kr.nama;
      if (statusP) muatPeriode(statusP.tanggal === statusP.tanggalMaster ? "" : statusP.tanggal);
      return bukaKoreksi(r.runId, nama).then(function () {
        $("kr-hasil").innerHTML = "";
        $("kr-hasil").appendChild(el("div", "info-box", "Tersimpan sebagai v" + r.versi + ". Total Net Flow " + rp(r.ringkasan.nf_total) +
          " · Migration " + rp(r.ringkasan.mig_total) + ". Konsolidasi, riwayat, dan jurnal periode ini sudah memakai angka baru."));
      });
    }).catch(function (e) {
      b.disabled = false;
      $("kr-hasil").innerHTML = "";
      $("kr-hasil").appendChild(el("div", "peringatan-box", e.message));
    });
  });

  // =====================================================================
  // Log proses (Tahap 5) — pengganti sheet "Audit Log"
  // File teks library\logs\proses\proses_{yyyy-MM}_{user}.txt, dibaca semua user per bulan
  // =====================================================================
  var dataLog = null, batasLog = 150, jalanLog = "";
  var STATUS_LOG = { OK: "OK", GAGAL: "Gagal", PERINGATAN: "Peringatan", INFO: "Info", DILEWATI: "Dilewati" };

  function muatLog(bulan) {
    teks("lp-info", "Memuat…");
    return panggil("logProses", { bulan: bulan || "" }).then(function (d) {
      dataLog = d;
      batasLog = 150;
      isiFilterLog();
      gambarLog();
    }).catch(function (e) { teks("lp-info", e.message); });
  }

  function labelBulan(b) { return BULAN_SINGKAT[+b.substr(5, 2) - 1] + " " + b.substr(0, 4); }

  function isiPilih(sel, nilai, labelSemua, fmtLabel) {
    var lama = sel.value;
    sel.innerHTML = "";
    if (labelSemua !== null) { var o = el("option", "", labelSemua); o.value = ""; sel.appendChild(o); }
    nilai.forEach(function (v) { var x = el("option", "", fmtLabel ? fmtLabel(v) : v); x.value = v; sel.appendChild(x); });
    sel.value = nilai.indexOf(lama) >= 0 ? lama : (labelSemua !== null ? "" : (nilai[0] || ""));
  }

  function isiFilterLog() {
    var bulan = dataLog.daftarBulan.slice();
    if (bulan.indexOf(dataLog.bulan) < 0) bulan.unshift(dataLog.bulan);
    isiPilih($("lp-bulan"), bulan, null, labelBulan);
    $("lp-bulan").value = dataLog.bulan;
    var proses = {}, user = {};
    dataLog.entri.forEach(function (e) { proses[e.proses] = true; user[e.pengguna] = true; });
    isiPilih($("lp-proses"), Object.keys(proses).sort(), "Semua");
    isiPilih($("lp-pengguna"), Object.keys(user).sort(), "Semua");
  }

  function rincianLog(r) {
    var hasil = [];
    (r || "").split("; ").forEach(function (b) {
      if (!b) return;
      var i = b.indexOf("=");
      hasil.push(i < 0 ? ["", b] : [b.substr(0, i), b.substr(i + 1)]);
    });
    return hasil;
  }

  function cocokLog(e) {
    var st = $("lp-status").value, pr = $("lp-proses").value, us = $("lp-pengguna").value;
    var cari = $("lp-cari").value.trim().toLowerCase();
    if (jalanLog && e.jalan !== jalanLog) return false;
    if (pr && e.proses !== pr) return false;
    if (us && e.pengguna !== us) return false;
    if (st === "masalah" && e.status !== "GAGAL" && e.status !== "PERINGATAN") return false;
    if (st && st !== "masalah" && e.status !== st) return false;
    if (cari) {
      var semua = [e.proses, e.label, e.status, e.sumber, e.pengguna, e.periode, e.jalan, e.rincian].join(" ").toLowerCase();
      if (semua.indexOf(cari) < 0) return false;
    }
    return true;
  }

  function gambarLog() {
    if (!dataLog) return;
    var list = dataLog.entri.filter(cocokLog);
    var ol = $("lp-daftar");
    ol.innerHTML = "";

    var info = dataLog.entri.length === 0
      ? "Belum ada catatan di " + NAMA_BULAN[+dataLog.bulan.substr(5, 2) - 1] + " " + dataLog.bulan.substr(0, 4) + "."
      : list.length + " dari " + dataLog.entri.length + " catatan · " + dataLog.file.length + " file · " +
        Math.max(1, Math.round(dataLog.ukuran / 1024)) + " KB";
    if (dataLog.terpotong) info += " · hanya 5.000 catatan terbaru yang dimuat";
    teks("lp-info", info);

    var bj = $("lp-jalan");
    bj.innerHTML = "";
    bj.hidden = !jalanLog;
    if (jalanLog) {
      bj.appendChild(el("span", "", "Hanya jalan " + jalanLog + " "));
      bj.appendChild(tombol("Tampilkan semua", "tautan", function () { jalanLog = ""; gambarLog(); }));
    }

    var hariLalu = "";
    list.slice(0, batasLog).forEach(function (e) {
      var hari = e.waktu.substr(0, 10);
      if (hari !== hariLalu) {
        hariLalu = hari;
        ol.appendChild(el("li", "lp-tanggal", tanggalPanjang(hari)));
      }
      ol.appendChild(barisLog(e));
    });
    $("lp-lagi").hidden = list.length <= batasLog;
  }

  function barisLog(e) {
    var li = el("li", "lp-item");
    var det = el("details");
    var sum = el("summary");
    sum.appendChild(el("span", "lp-waktu", e.waktu.substr(11, 8)));
    var judul = el("span", "lp-judul");
    judul.appendChild(el("span", "lp-proses", e.proses));
    if (e.label) judul.appendChild(el("span", "lp-label", " · " + e.label));
    judul.appendChild(el("span", "lp-sumber", e.sumber + " · " + e.pengguna));
    sum.appendChild(judul);
    var kelas = (e.status || "").toLowerCase();
    sum.appendChild(el("span", "lp-status " + kelas, STATUS_LOG[e.status] || e.status));
    det.appendChild(sum);

    var dl = el("dl", "detail lp-rinci");
    function baris(k, v) { dl.appendChild(el("dt", "", k)); var dd = el("dd", "", v); dl.appendChild(dd); return dd; }
    rincianLog(e.rincian).forEach(function (kv) {
      var dd = baris(kv[0] || "—", kv[1]);
      if (/^(file|snapshot|file awal|file akhir|file arsip)$/i.test(kv[0])) dd.className = "path";
    });
    if (e.periode) baris("Periode", tanggalPanjang(e.periode));
    if (e.jalan) {
      var dd = baris("Jalan", "");
      dd.appendChild(tombol(e.jalan + " — tampilkan semua langkahnya", "tautan", function () {
        jalanLog = e.jalan; batasLog = 500; gambarLog();
        $("lp-jalan").scrollIntoView({ block: "nearest" });
      }));
    }
    det.appendChild(dl);
    li.appendChild(det);
    return li;
  }

  function muatInfoAuditLog() {
    return panggil("infoAuditLog").then(function (d) {
      tampil("lp-lama", !!d.ada);
      if (d.ada) {
        teks("lp-lama-info", "Sheet Audit Log masih ada di workbook: " + fmt.format(d.baris) + " baris × " + d.kolom +
          " kolom. Modul yang sudah dipatch tidak menulis ke sheet ini lagi, jadi isinya bisa diarsipkan lalu sheet-nya dihapus.");
        tampil("lp-lama-arsip", true);
      }
    }).catch(function () { tampil("lp-lama", false); });
  }

  $("kartu-log").addEventListener("toggle", function () {
    if (!this.open) return;
    if (!dataLog) muatLog("");
    muatInfoAuditLog();
  });
  $("lp-muat").addEventListener("click", function () { muatLog($("lp-bulan").value); muatInfoAuditLog(); });
  $("lp-bulan").addEventListener("change", function () { jalanLog = ""; muatLog(this.value); });
  ["lp-proses", "lp-status", "lp-pengguna"].forEach(function (id) {
    $(id).addEventListener("change", function () { batasLog = 150; gambarLog(); });
  });
  $("lp-cari").addEventListener("input", function () { batasLog = 150; gambarLog(); });
  $("lp-lagi").addEventListener("click", function () { batasLog += 300; gambarLog(); });

  $("lp-lama-arsip").addEventListener("click", function () {
    tampil("lp-lama-konfirmasi", true);
    tampil("lp-lama-arsip", false);
    $("lp-lama-hasil").textContent = "";
  });
  $("lp-lama-batal").addEventListener("click", function () {
    tampil("lp-lama-konfirmasi", false);
    tampil("lp-lama-arsip", true);
  });
  $("lp-lama-ya").addEventListener("click", function () {
    var b = this;
    b.disabled = true;
    teks("lp-lama-hasil", "Mengarsipkan… (sheet besar bisa memerlukan beberapa menit)");
    panggil("arsipAuditLog", {}, 600000).then(function (r) {
      tampil("lp-lama-konfirmasi", false);
      if (!r.ada) { teks("lp-lama-hasil", "Sheet Audit Log sudah tidak ada."); return; }
      $("lp-lama-info").textContent = "";
      teks("lp-lama-hasil", fmt.format(r.baris) + " baris diarsipkan ke " + r.file +
        ". Sheet Audit Log sudah dihapus dari " + r.workbook + " — simpan workbook (Ctrl+S) agar ukuran file mengecil.");
      muatLog($("lp-bulan").value);
    }).catch(function (e) {
      teks("lp-lama-hasil", e.message);
      tampil("lp-lama-arsip", true);
      tampil("lp-lama-konfirmasi", false);
    }).then(function () { b.disabled = false; });
  });

  // =====================================================================
  // Diagnostik
  // =====================================================================
  function muatDiagnostik() {
    return panggil("ping").then(function (d) {
      infoPing = d;
      tampil("banner-koneksi", false);
      gambarPengaturan();
      $("d-jembatan").innerHTML = '<span class="ok">Terhubung</span>';
      teks("lencana-versi", "v" + d.versiAddin);
      teks("d-excel", d.versiExcel + " · " + d.arsitektur);
      teks("d-user", d.user);
      teks("d-hak", d.bolehMenulis ? "Lihat & simpan" : "Hanya lihat (nama user belum ada di config\\pengirim.txt)");
      teks("d-library", d.folderLibrary);
      teks("d-log", d.fileLog);
      teks("d-db", (d.databaseAda ? "Ada · " : "Belum ada (dibuat saat Simpan grup pertama) · ") + d.fileDatabase);
    }).catch(function (err) {
      $("d-jembatan").innerHTML = '<span class="err">Gagal</span> ' + err.message;
      $("diag").open = true;
      // Diagnostik kini ada di ⚙ Pengaturan, jadi kegagalan koneksi juga diumumkan di atas setiap tab
      var b = $("banner-koneksi");
      b.innerHTML = "";
      b.appendChild(document.createTextNode("Panel belum terhubung ke Excel: " + err.message + " "));
      b.appendChild(tombol("Lihat diagnostik", "tautan", function () {
        pindahTab("pengaturan");
        $("diag").scrollIntoView({ block: "start" });
      }));
      tampil("banner-koneksi", true);
    });
  }

  document.querySelectorAll("[data-folder]").forEach(function (b) {
    b.addEventListener("click", function () {
      panggil("bukaFolder", { folder: b.getAttribute("data-folder") }).catch(function () {});
    });
  });

  muatDiagnostik();
  muatPersiapan();
  muatPeriode();   // header (pemilih periode) & Ringkasan › Status periode
})();
