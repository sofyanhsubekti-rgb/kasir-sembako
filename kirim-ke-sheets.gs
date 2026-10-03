/**
 * Penerima data Kasir Sembako untuk Google Sheets.
 *
 * Cara pasang:
 *   1. Buat Google Sheet baru, beri nama misalnya "Laporan Kasir".
 *   2. Menu Ekstensi > Apps Script.
 *   3. Hapus isi Code.gs, tempel seluruh berkas ini, lalu Simpan.
 *   4. Ganti KODE_RAHASIA di bawah dengan kata sandi karangan sendiri.
 *   5. Klik Deploy > New deployment > pilih tipe "Web app".
 *        Execute as          : Me
 *        Who has access      : Anyone
 *      Klik Deploy, izinkan aksesnya, lalu salin "Web app URL".
 *   6. Tempel URL dan kode rahasia itu di aplikasi kasir:
 *      Pengaturan > Kirim ke Google Sheets.
 *
 * Kalau skripnya diubah, jangan lupa Deploy > Manage deployments >
 * ikon pensil > Version: New version > Deploy. Kalau tidak, yang jalan
 * masih versi lama.
 */

var KODE_RAHASIA = 'ganti-dengan-kata-sandi-anda';

/* Kolom tiap lembar. Kolom terakhir selalu ID, dipakai menolak kiriman ganda. */
var LEMBAR = {
  nota: {
    nama: 'Penjualan',
    judul: ['Tanggal', 'Jam', 'No Nota', 'Kasir', 'Metode', 'Total',
            'Modal (HPP)', 'Laba Kotor', 'Diskon', 'Rincian Barang', 'ID'],
    baris: function (d) {
      return [new Date(d.ts), jam_(d.ts), d.no || '', d.kasir || '', d.metode || '',
              angka_(d.total), angka_(d.hpp), angka_(d.total) - angka_(d.hpp),
              angka_(d.diskon), d.item || '', d.id];
    }
  },
  biaya: {
    nama: 'Pengeluaran',
    judul: ['Tanggal', 'Jam', 'Kategori', 'Keterangan', 'Sumber Dana', 'Nominal', 'ID'],
    baris: function (d) {
      return [new Date(d.ts), jam_(d.ts), d.kategori || '', d.ket || '',
              d.sumber === 'kas' ? 'Tunai' : 'Non-tunai', angka_(d.nominal), d.id];
    }
  },
  tutup: {
    nama: 'Tutup Kasir',
    judul: ['Tanggal', 'Shift', 'Kasir', 'Modal Awal', 'Tunai', 'Transfer', 'QRIS',
            'Omzet', 'Laba Kotor', 'Pengeluaran Tunai', 'Kas Seharusnya',
            'Uang Fisik', 'Selisih', 'Disetor', 'Catatan', 'ID'],
    baris: function (d) {
      return [new Date(d.ts), d.no || '', d.kasir || '', angka_(d.kasAwal),
              angka_(d.tunai), angka_(d.transfer), angka_(d.qris), angka_(d.omzet),
              angka_(d.omzet) - angka_(d.hpp), angka_(d.biayaKas),
              angka_(d.kasSeharusnya), angka_(d.fisik), angka_(d.selisih),
              angka_(d.setoran), d.catatan || '', d.id];
    }
  }
};

function doPost(e) {
  var kunci = LockService.getScriptLock();
  try {
    kunci.waitLock(25000);
  } catch (err) {
    return jawab_({ ok: false, pesan: 'server sibuk, coba lagi' });
  }
  try {
    var isi = JSON.parse(e.postData.contents);
    if (isi.kode !== KODE_RAHASIA) {
      return jawab_({ ok: false, pesan: 'kode rahasia salah' });
    }
    var daftar = isi.kirim || [];
    if (!daftar.length) return jawab_({ ok: true, diterima: [], pesan: 'tidak ada data' });

    var buku = SpreadsheetApp.getActiveSpreadsheet();
    var diterima = [];
    var kumpulan = {};

    /* Kelompokkan per jenis supaya satu lembar cukup sekali tulis. */
    daftar.forEach(function (d) {
      if (!d || !d.jenis || !LEMBAR[d.jenis] || !d.id) return;
      (kumpulan[d.jenis] = kumpulan[d.jenis] || []).push(d);
    });

    Object.keys(kumpulan).forEach(function (jenis) {
      var def = LEMBAR[jenis];
      var lembar = siapkanLembar_(buku, def);
      var sudahAda = idTersimpan_(lembar, def.judul.length);
      var barisBaru = [];

      kumpulan[jenis].forEach(function (d) {
        /* Kiriman ganda diabaikan: aplikasi boleh mengulang kirim
           tanpa takut data dobel. */
        if (sudahAda[d.id]) { diterima.push(d.id); return; }
        sudahAda[d.id] = true;
        barisBaru.push(def.baris(d));
        diterima.push(d.id);
      });

      if (barisBaru.length) {
        lembar.getRange(lembar.getLastRow() + 1, 1, barisBaru.length, def.judul.length)
              .setValues(barisBaru);
      }
    });

    perbaruiRingkasan_(buku);
    return jawab_({ ok: true, diterima: diterima });
  } catch (err) {
    return jawab_({ ok: false, pesan: String(err) });
  } finally {
    kunci.releaseLock();
  }
}

function doGet() {
  return jawab_({ ok: true, pesan: 'Penerima Kasir Sembako aktif.' });
}

function jawab_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}

function angka_(v) {
  var n = Number(v);
  return isFinite(n) ? n : 0;
}

function jam_(ts) {
  return Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'HH:mm');
}

function siapkanLembar_(buku, def) {
  var lembar = buku.getSheetByName(def.nama);
  if (!lembar) {
    lembar = buku.insertSheet(def.nama);
    lembar.appendRow(def.judul);
    lembar.getRange(1, 1, 1, def.judul.length).setFontWeight('bold');
    lembar.setFrozenRows(1);
    lembar.getRange('A:A').setNumberFormat('dd/MM/yyyy');
  }
  return lembar;
}

/** Kumpulkan ID yang sudah tercatat, supaya kiriman ulang tidak dobel. */
function idTersimpan_(lembar, jumlahKolom) {
  var baris = lembar.getLastRow();
  var ada = {};
  if (baris < 2) return ada;
  lembar.getRange(2, jumlahKolom, baris - 1, 1).getValues()
        .forEach(function (r) { if (r[0]) ada[String(r[0])] = true; });
  return ada;
}

/**
 * Spreadsheet berlokal Indonesia (dan banyak lokal Eropa) memakai titik-koma
 * sebagai pemisah argumen rumus; lokal Inggris memakai koma. Menebak dari nama
 * lokal gampang meleset, jadi dicoba langsung: tulis =SUM(1,1) lalu lihat
 * hasilnya 2 atau error.
 */
function pemisah_(buku) {
  var lembar = buku.getSheetByName('__cek__');
  if (!lembar) lembar = buku.insertSheet('__cek__');
  var sel = lembar.getRange('A1');
  sel.setFormula('=SUM(1,1)');
  SpreadsheetApp.flush();
  var komaJalan = sel.getValue() === 2;
  buku.deleteSheet(lembar);
  return komaJalan ? ',' : ';';
}

/** Lembar Ringkasan: yang dibuka pemilik toko dari HP. */
function perbaruiRingkasan_(buku) {
  var r = buku.getSheetByName('Ringkasan');
  if (r) return;                       // sudah ada, rumusnya hidup sendiri
  r = buku.insertSheet('Ringkasan', 0);

  var K = pemisah_(buku);              // pemisah argumen sesuai lokal
  var hariIni = 'Penjualan!A:A{K}">="&TODAY(){K}Penjualan!A:A{K}"<"&TODAY()+1';
  var biayaHariIni = 'Pengeluaran!A:A{K}">="&TODAY(){K}Pengeluaran!A:A{K}"<"&TODAY()+1';

  var isi = [
    ['RINGKASAN KASIR', ''],
    ['Diperbarui otomatis tiap ada transaksi masuk', ''],
    ['', ''],
    ['HARI INI', ''],
    ['Omzet', '=IFERROR(SUMIFS(Penjualan!F:F{K}' + hariIni + '){K}0)'],
    ['Jumlah nota', '=IFERROR(COUNTIFS(' + hariIni + '){K}0)'],
    ['Laba kotor', '=IFERROR(SUMIFS(Penjualan!H:H{K}' + hariIni + '){K}0)'],
    ['Pengeluaran', '=IFERROR(SUMIFS(Pengeluaran!F:F{K}' + biayaHariIni + '){K}0)'],
    ['Laba bersih', '=B7-B8'],
    ['', ''],
    ['7 HARI TERAKHIR', ''],
    ['Omzet', '=IFERROR(SUMIFS(Penjualan!F:F{K}Penjualan!A:A{K}">="&TODAY()-6){K}0)'],
    ['Laba kotor', '=IFERROR(SUMIFS(Penjualan!H:H{K}Penjualan!A:A{K}">="&TODAY()-6){K}0)'],
    ['Pengeluaran', '=IFERROR(SUMIFS(Pengeluaran!F:F{K}Pengeluaran!A:A{K}">="&TODAY()-6){K}0)'],
    ['Laba bersih', '=B13-B14'],
    ['', ''],
    ['SEPANJANG WAKTU', ''],
    ['Omzet', '=IFERROR(SUM(Penjualan!F:F){K}0)'],
    ['Laba kotor', '=IFERROR(SUM(Penjualan!H:H){K}0)'],
    ['Pengeluaran', '=IFERROR(SUM(Pengeluaran!F:F){K}0)'],
    ['Laba bersih', '=B19-B20'],
    ['', ''],
    ['OMZET PER HARI (14 hari)', ''],
    ['Tanggal', 'Omzet']
  ].map(function (baris) {
    return baris.map(function (sel) {
      return typeof sel === 'string' ? sel.split('{K}').join(K) : sel;
    });
  });
  r.getRange(1, 1, isi.length, 2).setValues(isi);

  /* Tabel omzet harian 14 hari ke belakang. */
  var mulai = isi.length + 1;
  for (var i = 0; i < 14; i++) {
    var b = mulai + i;
    r.getRange(b, 1).setFormula('=TODAY()-' + (13 - i));
    r.getRange(b, 2).setFormula(
      ('=IFERROR(SUMIFS(Penjualan!F:F{K}Penjualan!A:A{K}">="&A' + b +
       '{K}Penjualan!A:A{K}"<"&A' + b + '+1){K}0)').split('{K}').join(K));
  }

  r.getRange('A1').setFontSize(14).setFontWeight('bold');
  ['A4', 'A11', 'A17', 'A23'].forEach(function (sel) {
    r.getRange(sel).setFontWeight('bold');
  });
  r.getRange('A24:B24').setFontWeight('bold');
  r.getRange(mulai, 1, 14, 1).setNumberFormat('dd/MM/yyyy');
  r.getRange('B5:B21').setNumberFormat('#,##0');
  r.getRange(mulai, 2, 14, 1).setNumberFormat('#,##0');
  r.setColumnWidth(1, 230);
  r.setColumnWidth(2, 150);
}
