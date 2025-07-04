function updateLinkPDF() {
  const folderIterator = DriveApp.getFoldersByName("Your_Folder_Name");
  if (!folderIterator.hasNext()) {
    SpreadsheetApp.getUi().alert('Folder tidak ditemukan!');
    return;
  }

  const folder = folderIterator.next();
  const files = folder.getFiles();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Your_Sheet_Name");
  const data = sheet.getDataRange().getValues();

  // Buat map: nomor → baris
  const nomorMap = new Map();
  for (let i = 1; i < data.length; i++) {
    const nomor = data[i][1];
    const status = data[i][8]; // kolom ke-9 (I)

    // Lewati baris yang statusnya Sudah kembali
    if (status === "✅ Sudah Kembali") continue;

    nomorMap.set(nomor, i + 1); // baris di sheet (1-based)
  }

while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const link = file.getUrl();

    let start = null;
    let end = null;


    // 9. Pola: XXXX Teks Lain (Angka di awal bisa 1 atau lebih digit) - Paling umum, diletakkan terakhir
    const newSingleNumberPattern = fileName.match(/^(\d+)\s+.*$/i); // .txt opsional


   if (newSingleNumberPattern) {
      start = parseInt(newSingleNumberPattern[1], 10);
      end = start;
      Logger.log(`Matched newSingleNumberPattern: ${fileName} -> Start: ${start}, End: ${end}`);
    } else {
      // Tidak ada pola yang cocok
      Logger.log(`No match for file: ${fileName}`);
      start = null;
      end = null;
    }

    if (start !== null && end !== null) {
      const link = file.getUrl();

      for (let nomor = start; nomor <= end; nomor++) {
        if (nomorMap.has(nomor)) {
          const row = nomorMap.get(nomor);
          sheet.getRange(row, 5).setValue(fileName);
          sheet.getRange(row, 8).setValue(link);
          sheet.getRange(row, 9).setValue("📎 Terlampir");
        }
      }
    }
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Update link PDF selesai!');
  lastUpdate();
}

function lastUpdate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2025');
  sheet.getRange('H2').setValue('Last update: '+Utilities.formatDate(new Date(),"GMT+7","yyyy-MM-dd HH:mm:ss"))
}

function showBookingForm() {
 const html = HtmlService.createHtmlOutputFromFile("BookingUI")
   .setWidth(400)
   .setHeight(450);
 SpreadsheetApp.getUi().showModalDialog(html, "Form Pengambilan Nomor");
}

function processBooking(jenis, jumlah, pic) {
  var timeoutMs = 5000;  // wait max 5 seconds for the lock
  var lock = LockService.getScriptLock();
  var message = '';

  try {
    if (!lock.tryLock(timeoutMs)) {
      var msg = 'Script sedang berjalan. Coba lagi nanti';
      console.log(msg);
      try {
        SpreadsheetApp.getUi().alert(msg);
      } catch (e2) {
      }
      return;
    }

    if (jenis === "Automatic") {
      message = bookNumbersWithRange(jumlah, jenis, pic);
    } else {
      message = bookOthersWithRange(jumlah, jenis, pic);
    }

    // Tampilkan hasil booking ke user (kalau dipanggil via UI)
    try {
      SpreadsheetApp.getUi().alert("Booking berhasil:\n" + message);
    } catch (e3) {
      // Silent fail kalau bukan context UI
    }

  } catch (e) {
    console.error(e);
    try {
      SpreadsheetApp.getUi().alert("Error", e ? e.toString() : e, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e2) {
    }
  } finally {
    lock.releaseLock();
    SpreadsheetApp.flush();
  }
}

function bookNumbersWithRange(quantity, jenis, pic) {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const sheet = ss.getSheetByName('2025');

 const data = sheet.getDataRange().getValues();
 const today = new Date();
 let booked = 0;
 let startNomor = null;
 let endNomor = null;

 for (let i = 1; i < data.length && booked < quantity; i++) {
   if (data[i][8] === 'READY') {
     const nomor = parseInt(data[i][1]);
     if (booked === 0) startNomor = nomor;
     endNomor = nomor;

     sheet.getRange(i + 1, 3).setValue(pic);           // PIC
     sheet.getRange(i + 1, 6).setValue(today);         // Tanggal Booking
     sheet.getRange(i + 1, 7).setValue(jenis);         // Jenis
     sheet.getRange(i + 1, 9).setValue('🔒 Dipakai');  // Status

     booked++;
   }
 }

 if (booked < quantity) {
   throw new Error(`Nomor READY kurang. Hanya ${booked} yang berhasil di-book.`);
 }

 return `${String(startNomor).padStart(4, '0')} - ${String(endNomor).padStart(4, '0')}`;
}

function bookOthersWithRange(jumlah, jenis, pic) {
 const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2025");
 const data = sheet.getDataRange().getValues();

 const templateFile = DriveApp.getFileById("1c2eYVhvZqaRD7RzsOLTLkzfj4lJMlbkyKuN2o8Y83KY");
 const folderTujuan = DriveApp.getFoldersByName("File Excel Generate RI").next();

 // Cari nomor READY
 const readyRows = [];
 for (let i = 1; i < data.length; i++) {
   if (data[i][8] === "READY") readyRows.push({ nomor: parseInt(data[i][1]), row: i + 1 });
   if (readyRows.length >= jumlah) break;
 }

 if (readyRows.length < jumlah) {
   throw new Error(`Nomor READY kurang. Hanya ${readyRows.length} yang tersedia.`);
 }

 const today = new Date();
 const mulai = readyRows[0].nomor;
 const akhir = readyRows[jumlah - 1].nomor;

 for (let i = 0; i < jumlah; i++) {
   const { nomor, row } = readyRows[i];

   // Salin template dan set nomor
   const copy = templateFile.makeCopy(`Form serah Terima ${String(nomor).padStart(4, '0')}`, folderTujuan);
   const copySpreadsheet = SpreadsheetApp.openById(copy.getId());
   const sheetCopy = copySpreadsheet.getSheets()[0];
   sheetCopy.getRange("C4").setValue(`RI-00` + String(nomor).padStart(4, '0'));

   const url = copySpreadsheet.getUrl() + `#gid=${sheetCopy.getSheetId()}&range=C4`;

   // Update ke master sheet
   sheet.getRange(row, 3).setValue(pic);          // PIC
   sheet.getRange(row, 5).setValue(`Form serah Terima ${String(nomor).padStart(4, '0')}`); // Nama File
   sheet.getRange(row, 6).setValue(today);        // Tanggal
   sheet.getRange(row, 7).setValue(jenis);        // Jenis
   sheet.getRange(row, 8).setValue(url);          // Link
   sheet.getRange(row, 9).setValue("🔒 Dipakai");  // Status
 }

 return `${String(mulai).padStart(4, '0')} - ${String(akhir).padStart(4, '0')}`;
}

