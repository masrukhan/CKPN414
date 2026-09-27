// =====================================================================
// Panel CKPN — logika sisi panel (Tahap 1: fondasi)
//
// Komunikasi dengan C# (PanelBridge.cs):
//   kirim  : chrome.webview.postMessage({ id, cmd, args })
//   terima : { type: "reply", id, ok, data | error }
//            { type: "event", name, data }          (mulai Tahap 2)
// =====================================================================
(function () {
  "use strict";

  var webview = window.chrome && window.chrome.webview;
  var nextId = 1;
  var menunggu = {};              // id -> { resolve, reject, timer }
  var pendengarEvent = {};        // name -> [fn]

  // ---------------- Jembatan ----------------
  function panggil(cmd, args, batasMs) {
    return new Promise(function (resolve, reject) {
      if (!webview) { reject(new Error("Panel tidak berjalan di dalam Excel.")); return; }
      var id = nextId++;
      var timer = setTimeout(function () {
        delete menunggu[id];
        reject(new Error("Tidak ada jawaban dari library untuk '" + cmd + "'."));
      }, batasMs || 30000);
      menunggu[id] = { resolve: resolve, reject: reject, timer: timer };
      webview.postMessage({ id: id, cmd: cmd, args: args || {} });
    });
  }

  function onEvent(nama, fn) {
    (pendengarEvent[nama] = pendengarEvent[nama] || []).push(fn);
  }

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
        (pendengarEvent[m.name] || []).forEach(function (fn) { fn(m.data); });
      }
    });
  }

  // Disediakan global agar modul tahap berikutnya bisa memakai jembatan yang sama.
  window.CKPN = { panggil: panggil, onEvent: onEvent };

  // ---------------- Util ----------------
  function $(id) { return document.getElementById(id); }
  function teks(id, v) { $(id).textContent = (v === undefined || v === null || v === "") ? "—" : v; }

  // ---------------- Tab ----------------
  var tombolTab = document.querySelectorAll('.tab [role="tab"]');
  function pilihTab(nama) {
    tombolTab.forEach(function (b) {
      var aktif = b.getAttribute("data-tab") === nama;
      b.setAttribute("aria-selected", aktif ? "true" : "false");
      $("tab-" + b.getAttribute("data-tab")).hidden = !aktif;
    });
  }
  tombolTab.forEach(function (b) {
    b.addEventListener("click", function () { pilihTab(b.getAttribute("data-tab")); });
  });

  // ---------------- Diagnostik (ping) ----------------
  function muatDiagnostik() {
    return panggil("ping").then(function (d) {
      $("d-jembatan").innerHTML = '<span class="ok">Terhubung</span>';
      teks("lencana-versi", "v" + d.versiAddin);
      teks("d-excel", d.versiExcel + " · " + d.arsitektur);
      teks("d-webview", d.webview2);
      teks("d-user", d.user + " @ " + d.komputer);
      teks("d-library", d.folderLibrary);
      teks("d-db", d.databaseAda ? "ckpn.db ditemukan" : "Belum ada (dibuat di Tahap 3)");
    }).catch(function (err) {
      $("d-jembatan").innerHTML = '<span class="err">Gagal</span> ' + err.message;
      $("diag").open = true;
    });
  }

  // ---------------- Info workbook (baca Master) ----------------
  function muatWorkbook() {
    var status = $("wb-status");
    status.className = "nilai-besar";
    status.textContent = "Memeriksa…";
    return panggil("infoWorkbook").then(function (d) {
      if (!d.ditemukan) {
        status.className = "nilai-besar galat";
        status.textContent = d.pesan;
        $("wb-detail").hidden = true;
        return;
      }
      status.textContent = d.namaFile;
      $("wb-detail").hidden = false;
      teks("wb-bulan", d.bulanLaporan);
      teks("wb-topn", d.topN);
      teks("wb-kc", (d.kcDicentang || []).join(", "));
      teks("wb-file", d.fileIndividu);
    }).catch(function (err) {
      status.className = "nilai-besar galat";
      status.textContent = err.message;
    });
  }

  $("btn-muat-ulang").addEventListener("click", muatWorkbook);
  document.querySelectorAll("[data-folder]").forEach(function (b) {
    b.addEventListener("click", function () {
      panggil("bukaFolder", { folder: b.getAttribute("data-folder") }).catch(function () {});
    });
  });

  muatDiagnostik();
  muatWorkbook();
})();
