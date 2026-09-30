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
      if (nama === "riwayat") muatRiwayat();
      if (nama === "analisis") muatSumberAnalisis();
      Grafik.sembunyiTip();
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
          "Hitung dari tab Per grup untuk memakai Top-N tahunan secara otomatis."));
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

  // Status kartu grup (tab Per grup) disegarkan setiap kali isi workbook/database berubah:
  // selesai hitung, selesai simpan, dan saat kartu hasil/simpan ditutup — tanpa perlu pindah tab.
  function segarkanPerGrup() {
    if (!statusP) return;   // tab Per grup belum pernah dibuka; akan dimuat saat dibuka
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

  var NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  function tanggalPanjang(t) { return t ? (+t.substr(8, 2)) + " " + NAMA_BULAN[+t.substr(5, 2) - 1] + " " + t.substr(0, 4) : ""; }

  // =====================================================================
  // Tahap 4e: periode Master selalu terbaca di tab Per grup
  //   - header menampilkan Master!C4 dan tombol "Periksa ulang Master";
  //   - saat user kembali ke panel (fokus) setelah mengubah Master di Excel, periode & centang KC
  //     dicek; bila berubah, tab Per grup dan tab Manual dimuat ulang otomatis.
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
        if (!$("tab-periode").hidden || paksa) muatPeriode();   // "" = ikuti Master!C4
        else if (statusP) statusP = null;                       // dimuat ulang saat tab dibuka
        muatPersiapan();                                        // tab Manual ikut segar
      }
    }).catch(function () {}).then(function () { sedangCekMaster = false; });
  }
  $("btn-periksa-master").addEventListener("click", function () { cekMaster(true); });
  window.addEventListener("focus", function () { cekMaster(false); });
  document.addEventListener("visibilitychange", function () { if (!document.hidden) cekMaster(false); });

  function gambarAlur(d) {
    teks("pg-master", d.tanggalMaster ? tanggalPanjang(d.tanggalMaster) : "Master!C4 belum diisi");
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
    var bolehSemua = bolehHitung && d.bolehMenulis && d.susunanDitetapkan;
    $("btn-batch-opsi").disabled = !bolehSemua;
    $("btn-batch-opsi").title = bolehSemua ? "" : "Tersedia untuk periode Master yang belum Final, setelah susunan grup ditetapkan, bagi user pengirim.";
    if (!bolehSemua) tampil("kartu-batch-opsi", false);

    (d.grup || []).forEach(function (g) {
      var k = el("div", "kartu-grup" + (g.diWorkbook ? " di-workbook" : ""));
      var head = el("div", "kepala-grup");
      var kiri = el("div");
      kiri.appendChild(el("div", "nama-grup", g.nama));
      kiri.appendChild(el("div", "teks-kecil", g.kc.join(", ") + (g.topN ? " · Top-N " + g.topN : "") + (g.diWorkbook ? " · sedang di workbook" : "")));
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
  });
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
  var LABEL_BATCH = {
    "antre": "menunggu", "menghitung": "menghitung…", "tersimpan": "tersimpan", "perlu-review": "perlu review",
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
        "Grup tanpa temuan langsung disimpan sebagai versi baru.")) return;
    b.disabled = true;
    panggil("hitungSemuaGrup", { lewatiTersimpan: $("batch-lewati").checked, berhentiBilaTemuan: berhenti })
      .then(function () { tampil("kartu-batch-opsi", false); window.scrollTo(0, 0); })
      .catch(function (e) { alert(e.message); })
      .then(function () { b.disabled = false; });
  });
  $("btn-batch-tutup").addEventListener("click", function () { tampil("kartu-batch", false); });

  onEvent("batchMulai", function (d) {
    batchAktif = true;
    runTerakhirBatch = null;
    tampil("kartu-hasil", false);
    tampil("kartu-simpan", false);
    tampil("btn-batch-tutup", false);
    teks("batch-judul", "Hitung semua grup · " + d.tanggal);
    teks("batch-sub", d.grup.length + " grup · " + (d.berhentiBilaTemuan ? "berhenti bila ada temuan" : "tandai temuan lalu lanjut"));
    var ol = $("batch-daftar");
    ol.innerHTML = "";
    d.grup.forEach(function (g, i) {
      var li = el("li", "antre");
      li.id = "bg-" + i;
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
    var chip = li.querySelector(".status-grup");
    chip.className = "status-grup " + (d.status === "tersimpan" ? "tersimpan" : d.status === "perlu-review" || d.status === "menghitung" ? "dihitung" : "");
    chip.textContent = LABEL_BATCH[d.status] + (d.versi ? " v" + d.versi : "");
    var ul = li.querySelector(".batch-temuan");
    ul.innerHTML = "";
    (d.temuan || []).forEach(function (t) { ul.appendChild(el("li", "", t)); });
    if (d.status === "tersimpan" && (d.nfTotal !== undefined || d.migTotal !== undefined))
      ul.appendChild(el("li", "ok-ringan", "Net Flow " + rp(d.nfTotal) + " · Migration " + rp(d.migTotal)));
  });

  onEvent("batchSelesai", function (d) {
    batchAktif = false;
    var n = { tersimpan: 0, review: 0 };
    (d.hasil || []).forEach(function (h) { if (h.status === "tersimpan") n.tersimpan++; if (h.status === "perlu-review") n.review++; });
    var w = $("batch-ringkas");
    w.innerHTML = "";
    var teksRingkas = n.tersimpan + " grup tersimpan otomatis" + (n.review ? " · " + n.review + " perlu review" : "") +
      (d.sisa ? " · " + d.sisa + " grup belum dihitung" : "");
    if (d.status === "selesai") w.appendChild(el("div", "status-hasil ok", "Selesai · " + teksRingkas));
    else if (d.status === "perlu-review") {
      w.appendChild(el("div", "status-hasil batal", "Berhenti untuk review · " + teksRingkas));
      w.appendChild(el("div", "teks-kecil", "Sheet berisi grup yang perlu direview. Cek temuan di bawah, edit bila perlu, lalu simpan. " +
        "Setelah itu jalankan lagi Hitung semua grup (grup tersimpan dilewati) untuk melanjutkan."));
    } else if (d.status === "dibatalkan") w.appendChild(el("div", "status-hasil batal", "Dibatalkan · " + teksRingkas));
    else w.appendChild(el("div", "status-hasil err", "Gagal: " + (d.error || "") + " · " + teksRingkas));
    if (n.review && d.status !== "perlu-review")
      w.appendChild(el("div", "teks-kecil", "Grup bertanda perlu review: hitung satu per satu dari kartu grupnya (Hitung grup ini), review, lalu simpan."));
    tampil("btn-batch-tutup", true);
    // Grup tempat proses berhenti: tampilkan kartu hasil + review agar bisa langsung disimpan
    if (d.status === "perlu-review" && runTerakhirBatch) { tampilkanHasil(runTerakhirBatch); muatReview(); }
    else tampil("kartu-langkah", true);
    muatPersiapan();
    segarkanPerGrup();
  });

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
      var tr = el("tr");
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
    if (!grup) $("rw-tabel").appendChild(el("p", "teks-kecil", "Kolom yang disorot = metode terpakai. Selisih positif berarti PPKA lebih kecil dari CKPN."));
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
      if (!$("tab-riwayat").hidden && dataRw) gambarRiwayat();
      if (!$("tab-analisis").hidden && dataAn) { if (dataAn.migrasi) siapkanMigrasi(); if (dataAn.netflow) gambarNetFlow(); }
    }, 200);
  });

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
