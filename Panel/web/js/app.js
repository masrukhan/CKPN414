// =====================================================================
// Panel CKPN — logika sisi panel (Tahap 2: menjalankan perhitungan)
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
      if (nama === "periode") { muatStaging(); muatPenyesuaian(); }
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
    // tab Jalankan harus aktif agar kartu terlihat
    document.querySelector('.tab [data-tab="jalankan"]').click();
    periksaSimpan();
  }

  function periksaSimpan() {
    $("btn-simpan").disabled = true;
    teks("simpan-info", "Memeriksa sheet…");
    $("simpan-peringatan").innerHTML = "";
    $("simpan-perubahan").innerHTML = "";
    panggil("periksaSimpan").then(function (d) {
      dataSimpan = d;
      teks("simpan-info", d.periode + " · " + d.kodeKC + " · akan disimpan sebagai versi " + d.versi +
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
        if (x.perluAlasan) {
          var sel = el("select", "isian");
          sel.setAttribute("aria-label", "Alasan penghapusan " + x.kunci);
          sel.dataset.kunci = x.kunci;
          sel.className = "isian alasan-hapus";
          var kosong = el("option", "", "— pilih alasan (wajib) —"); kosong.value = "";
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

  function kumpulkanAlasan() {
    var hasil = {}, kurang = [];
    document.querySelectorAll("select.alasan-hapus").forEach(function (sel) {
      var v = sel.value === "__lain" ? sel.nextSibling.value.trim() : sel.value;
      if (!v) kurang.push(sel.dataset.kunci); else hasil[sel.dataset.kunci] = v;
    });
    return { alasan: hasil, kurang: kurang };
  }

  $("btn-simpan").addEventListener("click", function () {
    var a = kumpulkanAlasan();
    if (a.kurang.length) {
      $("simpan-hasil").innerHTML = "";
      $("simpan-hasil").appendChild(el("div", "peringatan-box", "Alasan belum diisi untuk: " + a.kurang.join(", ")));
      return;
    }
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
  $("btn-simpan-dari-periode").addEventListener("click", bukaSimpan);
  $("btn-simpan-periksa").addEventListener("click", periksaSimpan);
  $("btn-simpan-batal").addEventListener("click", function () {
    tampil("kartu-simpan", false);
    tampil("kartu-langkah", true);
  });

  // =====================================================================
  // 7. Tab Periode: kiriman grup & penyesuaian tersimpan
  // =====================================================================
  function muatStaging(tanggal) {
    teks("staging-isi", "Memuat…");
    panggil("daftarStaging", { tanggal: tanggal || "" }).then(function (d) {
      var w = $("staging-isi");
      w.innerHTML = "";
      if (!d.adaDatabase) { w.textContent = "Database belum ada. Simpan grup pertama akan membuatnya."; return; }

      var sel = $("pilih-periode");
      sel.innerHTML = "";
      (d.periode || []).forEach(function (p) {
        var o = el("option", "", p.tanggal + " · " + p.jumlahGrup + " grup");
        o.value = p.tanggal;
        if (p.tanggal === d.tanggal) o.selected = true;
        sel.appendChild(o);
      });

      var runs = d.runs || [];
      if (!runs.length) { w.textContent = "Belum ada grup yang disimpan untuk periode ini."; return; }
      var t = el("table", "tabel-staging");
      var h = el("tr");
      ["Grup (KC)", "Versi", "Net Flow", "Migration", "PPKA", ""].forEach(function (x) { h.appendChild(el("th", "", x)); });
      t.appendChild(h);
      runs.forEach(function (r) {
        var tr = el("tr", r.aktif ? "" : "lama");
        var c1 = el("td");
        c1.appendChild(el("div", "", r.kodeKC.replace(/,/g, ", ")));
        c1.appendChild(el("div", "teks-kecil", r.pengguna + " · " + r.waktu));
        tr.appendChild(c1);
        var c2 = el("td", "", "v" + r.versi + " ");
        if (r.aktif) c2.appendChild(el("span", "lencana-aktif", "aktif"));
        tr.appendChild(c2);
        var rg = r.ringkasan || {};
        tr.appendChild(el("td", "angka", rp(rg.nf_total)));
        tr.appendChild(el("td", "angka", rp(rg.mig_total)));
        tr.appendChild(el("td", "angka", rp(rg.ppka_total)));
        var c6 = el("td");
        if (r.adaSnapshot) {
          var b = el("button", "tautan", "xlsx");
          b.type = "button";
          b.title = "Buka snapshot";
          b.addEventListener("click", function () { panggil("bukaSnapshot", { runId: r.id }).catch(function (e) { alert(e.message); }); });
          c6.appendChild(b);
        }
        tr.appendChild(c6);
        t.appendChild(tr);
      });
      w.appendChild(t);
      w.appendChild(el("p", "teks-kecil", "Nilai PPKA dan CKPN diambil dari sheet Summary saat grup disimpan."));
    }).catch(function (e) { teks("staging-isi", e.message); });
  }

  $("pilih-periode").addEventListener("change", function () { muatStaging(this.value); });

  function muatPenyesuaian() {
    teks("penyesuaian-isi", "Memuat…");
    panggil("daftarPenyesuaian").then(function (d) {
      var w = $("penyesuaian-isi");
      w.innerHTML = "";
      var ind = d.individu || [], cs = d.lgdcs || [];
      if (!ind.length && !cs.length) { w.textContent = "Belum ada penyesuaian tersimpan."; return; }

      function item(modul, kunci, uraian, info) {
        var row = el("div", "item-penyesuaian");
        var kiri = el("div");
        kiri.appendChild(el("span", "kunci", kunci));
        kiri.appendChild(el("div", "teks-kecil", uraian));
        kiri.appendChild(el("div", "teks-kecil", info));
        row.appendChild(kiri);
        if (d.bolehMenulis) {
          var b = el("button", "", "Hapus");
          b.type = "button";
          b.addEventListener("click", function () {
            var alasan = window.prompt("Alasan menghapus penyesuaian " + kunci + ":");
            if (!alasan) return;
            panggil("hapusPenyesuaian", { modul: modul, kunci: kunci, alasan: alasan })
              .then(muatPenyesuaian).catch(function (e) { alert(e.message); });
          });
          row.appendChild(b);
        }
        return row;
      }

      if (ind.length) {
        w.appendChild(el("h4", "", "CKPN Individu (" + ind.length + ")"));
        ind.forEach(function (p) {
          w.appendChild(item("individu", p.kunci,
            "Agunan " + rp(p.jaminan) + " (sistem " + rp(p.jaminanSistem) + ") · biaya " + rp(p.biaya),
            p.pengguna + " · " + p.waktu + (p.alasan ? " · " + p.alasan : "")));
        });
      }
      if (cs.length) {
        w.appendChild(el("h4", "", "LGD Collateral Shortfall (" + cs.length + ")"));
        cs.forEach(function (p) {
          var u = p.jenis === "hapus" ? "Dikecualikan" :
                  p.jenis === "tambah" ? "Baris manual · pokok " + rp(p.pokok) + " · realisasi " + rp(p.recovery) :
                  "Diubah" + (p.nilaiAgunan !== null ? " · agunan " + rp(p.nilaiAgunan) : "") +
                             (p.recovery !== null ? " · realisasi " + rp(p.recovery) : "");
          w.appendChild(item("lgdcs", p.kunci, (p.nama ? p.nama + " · " : "") + u,
            p.pengguna + " · " + p.waktu + (p.alasan ? " · " + p.alasan : "")));
        });
      }
    }).catch(function (e) { teks("penyesuaian-isi", e.message); });
  }
  $("btn-muat-penyesuaian").addEventListener("click", muatPenyesuaian);

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
