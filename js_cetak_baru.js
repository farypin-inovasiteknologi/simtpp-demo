function showDownloadModal(url) { let modalEl = document.getElementById('modalDownloadReady'); if(modalEl) { document.getElementById('btnRealDownload').href = url; let modalObj = bootstrap.Modal.getInstance(modalEl) || new bootstrap.Modal(modalEl); modalObj.show(); } else { window.open(url, '_blank'); } }
  function tutupModalDownload() { setTimeout(() => { let modalEl = document.getElementById('modalDownloadReady'); let modal = bootstrap.Modal.getInstance(modalEl); if(modal) modal.hide(); }, 1000); }

  function cetakPDFPerorangan(tabId, namaFile) { window.scrollTo(0, 0); let element = document.getElementById('viewManajemenASN'); let navPills = element.querySelector('.nav-pills'); let actionBtns = element.querySelectorAll('button'); let origPillsDisplay = navPills ? navPills.style.display : ''; if(navPills) navPills.style.display = 'none'; actionBtns.forEach(b => { b.dataset.origDisplay = b.style.display; b.style.display = 'none'; }); let tables = element.querySelectorAll('.table-responsive'); tables.forEach(t => { t.style.overflow = 'visible'; }); let selects = element.querySelectorAll('.form-select'); selects.forEach(s => { s.dataset.origBg = s.style.backgroundImage; s.style.backgroundImage = 'none'; }); startLoading("Menyusun PDF..."); let opt = { margin: 0.3, filename: `${namaFile}_${asNIPAktif}.pdf`, image: { type: 'jpeg', quality: 0.98 }, html2canvas: { scale: 2, useCORS: true, scrollY: 0, scrollX: 0 }, jsPDF: { unit: 'in', format: 'legal', orientation: 'portrait' } }; html2pdf().set(opt).from(element).save().then(() => { stopLoading(); if(navPills) navPills.style.display = origPillsDisplay; actionBtns.forEach(b => { b.style.display = b.dataset.origDisplay; }); tables.forEach(t => { t.style.overflow = ''; }); selects.forEach(s => { s.style.backgroundImage = s.dataset.origBg; }); setTimeout(() => { alertSukses("File PDF Perorangan berhasil diunduh!"); }, 500); }).catch(err => { stopLoading(); if(navPills) navPills.style.display = origPillsDisplay; actionBtns.forEach(b => { b.style.display = b.dataset.origDisplay; }); tables.forEach(t => { t.style.overflow = ''; }); selects.forEach(s => { s.style.backgroundImage = s.dataset.origBg; }); setTimeout(() => { alertError("Gagal memproses PDF: " + err); }, 500); }); }

  function cetakTTD(sheet, startRow, setting, colMax = null) {
      const letRight = colMax ? numToLet(colMax - 1) : 'P';
      sheet.getCell(`${letRight}${startRow}`).value = `Jambi, ........................ ${new Date().getFullYear()}`;
      sheet.getCell(`B${startRow+1}`).value = "Mengetahui,";
      sheet.getCell(`B${startRow+2}`).value = setting.Kepala_Jabatan;
      sheet.getCell(`${letRight}${startRow+2}`).value = "BENDAHARA PENGELUARAN";
      
      sheet.getCell(`B${startRow+6}`).value = setting.Kepala_Nama;
      sheet.getCell(`B${startRow+6}`).font = { bold: true, underline: true };
      sheet.getCell(`${letRight}${startRow+6}`).value = setting.Bendahara_Nama;
      sheet.getCell(`${letRight}${startRow+6}`).font = { bold: true, underline: true };
      
      sheet.getCell(`B${startRow+7}`).value = setting.Kepala_Pangkat;
      sheet.getCell(`${letRight}${startRow+7}`).value = setting.Bendahara_Pangkat;
      
      sheet.getCell(`B${startRow+8}`).value = "NIP. " + setting.Kepala_NIP;
      sheet.getCell(`${letRight}${startRow+8}`).value = "NIP. " + setting.Bendahara_NIP;
  }

  function unduhFile(workbook, namaFile) {
      workbook.xlsx.writeBuffer().then((buffer) => {
          let blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
          saveAs(blob, namaFile);
          alertSukses("Laporan Excel berhasil diunduh!");
      }).catch(err => { alertError("Gagal men-generate Excel: " + err); });
  }

  // ==============================================================
  // UI COUNTER & PROGRESS BAR
  // ==============================================================
  function showLoadingPhase1() {
      Swal.fire({ 
          title: 'Menyiapkan Laporan...', 
          html: `<div id="loadStatus" class="fw-bold text-primary mb-2 fs-6">1. Menarik data dari Server...</div>
                 <div id="boxProgress" class="hidden mt-3">
                    <h5 class="text-success fw-bold" id="progressCounter">Memproses: 0 / 0</h5>
                    <div class="progress shadow-sm" style="height: 18px; border-radius: 10px; background-color: #e9ecef;">
                        <div id="progressBar" class="progress-bar progress-bar-striped progress-bar-animated bg-success" style="width: 0%"></div>
                    </div>
                 </div>`, 
          allowOutsideClick: false, showConfirmButton: false, didOpen: () => { Swal.showLoading(); } 
      });
  }

  function switchToPhase2(totalData) {
      let statusEl = document.getElementById('loadStatus');
      if(statusEl) {
          statusEl.innerText = "2. Menyusun File Excel...";
          statusEl.className = "fw-bold text-success mb-2 fs-6";
          document.getElementById('boxProgress').classList.remove('hidden');
          document.getElementById('progressCounter').innerText = `Memproses: 0 / ${totalData}`;
      }
  }

  function updateProgress(current, total) {
      let counterEl = document.getElementById('progressCounter');
      let barEl = document.getElementById('progressBar');
      if(counterEl) counterEl.innerText = `Memproses: ${current} / ${total}`;
      if(barEl) barEl.style.width = `${Math.round((current / total) * 100)}%`;
  }


  // ==============================================================
  // TRIGGER DOWNLOAD (SUDAH DIPERBAIKI ANTI-FREEZE)
  // ==============================================================
  async function fetchDataLaporan(format, e, funcExcel, funcPdf) {
      if(e) e.preventDefault(); 
      if(!globalBulanAktif) return alertPeringatan("Pilih bulan terlebih dahulu!"); 
      
      // 1. BLOKIR CETAK PDF MASSAL DARI AWAL
      if (format.toLowerCase() === 'pdf') {
          return Swal.fire({
              icon: 'info',
              title: 'PDF Massal Dinonaktifkan',
              text: 'Format tabel ini terlalu lebar (banyak kolom) untuk dijadikan PDF secara otomatis. Silakan unduh format Excel, lalu "Save As PDF" secara manual melalui aplikasi Microsoft Excel di laptop Anda.',
              confirmButtonColor: '#0d6efd'
          });
      }

      let fUnit = document.getElementById('filterUnitKerja').value; 
      
      showLoadingPhase1(); 
      try {
          let actionName = (funcExcel === buatExcelNominatifJS) ? "getJsonNominatif" : "getJsonLaporanLengkap";
          let payload = { bulanAktif: globalBulanAktif, jenisASN: globalJenisASN, roleUser: currentUser.role, unitkerjaUser: currentUser.unitkerja, filterUnit: fUnit };
          
          let res = await fetchAPI(actionName, payload); 
          
          if(res && res.error) {
              Swal.close(); return alertError("❌ " + res.error); 
          } else if (res && res.status === "sukses") {
              
              // Jeda sejenak agar browser sempat mempersiapkan UI
              await new Promise(r => setTimeout(r, 100));

              let adaBelumSimpan = res.data.some(c => (c.gajiKotor === 0 || c.gajiKotorTER === 0));
              
              if (adaBelumSimpan) {
                  Swal.close(); 
                  Swal.fire({
                      title: 'Ada Data Belum Lengkap!',
                      text: 'Ditemukan pegawai yang data Gaji-nya belum tersimpan (angkanya Rp 0). Yakin ingin lanjut mencetak Excel?',
                      icon: 'warning', showCancelButton: true, confirmButtonColor: '#3085d6', cancelButtonColor: '#d33',
                      confirmButtonText: 'Ya, Lanjut Cetak!', cancelButtonText: 'Batal'
                  }).then(async (result) => {
                      if (result.isConfirmed) {
                          showLoadingPhase1(); 
                          switchToPhase2(res.data.length); 
                          await new Promise(r => setTimeout(r, 100)); // Napas buatan untuk UI
                          
                          try {
                              await funcExcel(res); 
                          } catch(excelErr) {
                              Swal.close(); alertError("Terjadi error sistem saat menyusun Excel: " + excelErr.message);
                          }
                      }
                  });
              } else {
                  switchToPhase2(res.data.length); 
                  await new Promise(r => setTimeout(r, 100)); // Napas buatan untuk UI
                  try {
                      await funcExcel(res); 
                  } catch(excelErr) {
                      Swal.close(); alertError("Terjadi error sistem saat menyusun Excel: " + excelErr.message);
                  }
              }
          } else { Swal.close(); alertError("❌ Gagal mengambil data dari server. Periksa koneksi."); } 
      } catch (err) { Swal.close(); alertError("Terjadi kesalahan sistem: " + err.message); }
  }

  function unduhNominatifKolektif(format, e) { fetchDataLaporan(format, e, buatExcelNominatifJS, null); } 
  function unduhPerhitunganKolektif(format, e) { fetchDataLaporan(format, e, buatExcelPerhitunganJS, null); }
  function unduhRekapGolongan(format, e) { fetchDataLaporan(format, e, buatExcelRekapGolonganJS, null); }
  function unduhRekening(format, e) { fetchDataLaporan(format, e, buatExcelRekeningJS, null); }
  function unduhRekapPajak(format, e) { fetchDataLaporan(format, e, buatExcelRekapPajakJS, null); }


  // ==============================================================
  // 1. ENGINE EXCELJS: DAFTAR NOMINATIF (18 KOLOM - BEBAS ERROR)
  // ==============================================================
  async function buatExcelNominatifJS(res) {
      try {
          const wb = new ExcelJS.Workbook();
          const sheet = wb.addWorksheet('Daftar Nominatif', { pageSetup: { paperSize: 5, orientation: 'landscape', margins: { left: 0.2, right: 0.2, top: 0.4, bottom: 0.4 } } });

          const colWidths = [5, 38, 15, 30, 15, 11, 11, 11, 11, 11, 14, 12, 15, 12, 12, 12, 14, 18];
          colWidths.forEach((w, i) => { sheet.getColumn(i+1).width = w; });

          sheet.mergeCells('A1:R1'); sheet.getCell('A1').value = `DAFTAR NOMINATIF TPP ${res.setting.Nama_Dinas} PEMERINTAH PROVINSI JAMBI`;
          sheet.mergeCells('A2:R2'); sheet.getCell('A2').value = `${res.jenisASN} ${res.unitCetak}`;
          sheet.mergeCells('A3:R3'); sheet.getCell('A3').value = `PERIODE BULAN: ${res.bulanBesar}`;
          
          for(let i=1; i<=3; i++) { 
              let c = sheet.getCell(`A${i}`);
              c.font = { bold: true, size: i===1 ? 14 : 12 }; 
              c.alignment = { horizontal: 'center', vertical: 'middle' }; 
          }

          sheet.mergeCells('A5:A6'); sheet.getCell('A5').value = "No.";
          sheet.mergeCells('B5:B6'); sheet.getCell('B5').value = "Nama / Tgl Lahir / NIP / Gol.";
          sheet.mergeCells('C5:C6'); sheet.getCell('C5').value = "Status / Jiwa";
          sheet.mergeCells('D5:D6'); sheet.getCell('D5').value = "Jabatan";
          sheet.mergeCells('E5:E6'); sheet.getCell('E5').value = "Gaji Kotor";
          
          sheet.mergeCells('F5:M5'); sheet.getCell('F5').value = "PERHITUNGAN TPP (NETTO)";
          ['BK','PK','KK','TB','KP','Jumlah TPP','BPJS 4%','Total TPP Netto'].forEach((txt, i) => sheet.getCell(6, 6+i).value = txt);
          
          sheet.mergeCells('N5:Q5'); sheet.getCell('N5').value = "PENGURANGAN TPP";
          ['IWP 1%','PPh 21','BPJS 4%','TOTAL POT.'].forEach((txt, i) => sheet.getCell(6, 14+i).value = txt);
          
          sheet.mergeCells('R5:R6'); sheet.getCell('R5').value = "TPP BERSIH DITERIMA";

          for(let i=1; i<=18; i++) { sheet.getCell(7, i).value = i.toString(); }

          for (let r = 5; r <= 7; r++) {
              let row = sheet.getRow(r);
              row.height = r === 7 ? 15 : 25;
              for(let c = 1; c <= 18; c++) {
                  let cell = row.getCell(c);
                  cell.font = { bold: r!==7, italic: r===7, size: r===7 ? 8 : 10 };
                  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                  cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                  
                  let bg = 'FFCFE2F3'; // Kode Warna ARGB Wajib 8 Digit!
                  if(r === 7) bg = 'FFE9ECEF'; 
                  else if (c >= 14 && c <= 17 && r === 5) bg = 'FFFCE5CD'; 
                  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
              }
          }

          let grouped = {};
          res.data.forEach(c => {
              let golDasar = c.golonganAsli.split("/")[0].trim().toUpperCase();
              let groupName = (res.jenisASN === "PPPK") ? getGroupPPPK(golDasar) : (golDasar.includes("IX") || golDasar.includes("X") ? "PPPK" : golDasar);
              if(!grouped[groupName]) grouped[groupName] = [];
              grouped[groupName].push(c);
          });

          let urutanGol = res.jenisASN === "PPPK" ? ["XVII - XIII", "XII - IX", "VIII - V", "IV - I"] : ["IV", "III", "II", "I"];
          let noUrut = 1; let currentRow = 8;
          let progressCount = 0; let totalData = res.data.length;
          
          // Penampung Nilai Kalkulator JS
          let grandTotals = Array(14).fill(0); 

          for (let gol of urutanGol) {
              if(grouped[gol] && grouped[gol].length > 0) {
                  let subTotals = Array(14).fill(0); // Penampung Sub Total JS

                  for (let c of grouped[gol]) {
                      let isKawin = c.statusTER.startsWith("K");
                      let jmlJiwa = (isKawin ? 2 : 1) + parseInt(c.tanggungAnak || 0);
                      let rasio = c.tppBruto > 0 ? (c.tppNettoKinerja / c.tppBruto) : 0;
                      let tppPlusBpjs = c.tppNettoKinerja + c.bpjs4;
                      let totalPot = c.iwp1 + c.pph21TKD + c.bpjs4; 
                      let tppBersih = tppPlusBpjs - totalPot;

                      // Array nilai uang
                      let vals = [
                          c.gajiKotor, Math.round(c.bk * rasio), Math.round(c.pk * rasio), Math.round(c.kk * rasio), Math.round(c.tb * rasio), Math.round(c.kp * rasio),
                          c.tppNettoKinerja, c.bpjs4, tppPlusBpjs, c.iwp1, c.pph21TKD, c.bpjs4, totalPot, tppBersih
                      ];

                      // Tambahkan ke SubTotal dan GrandTotal JS
                      vals.forEach((v, i) => { subTotals[i] += v; grandTotals[i] += v; });

                      let row = sheet.getRow(currentRow);
                      row.height = 70;
                      row.values = [ noUrut++, `${c.nama}\n${c.tglLahir}\nNIP. ${c.nip}\n${res.jenisASN} - Gol. ${c.golonganAsli}`, `${c.statusTER}\nJiwa: ${jmlJiwa}`, c.jabatan, ...vals ];

                      row.eachCell({ includeEmpty: true }, (cell, colN) => {
                          cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                          cell.alignment = { vertical: 'middle', wrapText: true, horizontal: colN===2 || colN===4 ? 'left' : (colN>4 ? 'right' : 'center') };
                          if (colN >= 5) cell.numFmt = '#,##0';
                          if (cell.value === undefined && colN >= 5) cell.value = 0;
                      });
                      currentRow++; progressCount++;
                      
                      updateProgress(progressCount, totalData); 
                      if (progressCount % 5 === 0 || progressCount === totalData) { await new Promise(r => setTimeout(r, 0)); }
                  }

                  let rowSub = sheet.getRow(currentRow);
                  rowSub.height = 25; 
                  // Tulis langsung angka SubTotal ke cell (TANPA RUMUS)
                  rowSub.values = ["", `SUB-TOTAL GOLONGAN ${gol}`, "", "", ...subTotals];
                  rowSub.getCell(2).font = { bold: true };

                  rowSub.eachCell({ includeEmpty: true }, (cell, colN) => {
                      cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F3F3' } };
                      if (colN >= 5) { cell.numFmt = '#,##0'; cell.font = {bold: true}; }
                  });
                  currentRow++;
              }
          }

          let rowGrand = sheet.getRow(currentRow);
          rowGrand.height = 30; 
          // Tulis langsung angka Grand Total ke cell (TANPA RUMUS)
          rowGrand.values = ["", "TOTAL KESELURUHAN (ALL GOLONGAN)", "", "", ...grandTotals];
          rowGrand.getCell(2).font = { bold: true };

          rowGrand.eachCell({ includeEmpty: true }, (cell, colN) => {
              cell.border = { top: {style:'medium'}, left: {style:'thin'}, bottom: {style:'medium'}, right: {style:'thin'} };
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9ECEF' } };
              if (colN >= 5) { cell.numFmt = '#,##0'; cell.font = {bold: true}; }
          });

          cetakTTD(sheet, currentRow + 3, res.setting, 18);
          unduhFile(wb, `Daftar_Nominatif_${res.bulanBesar}_${res.unitCetak}.xlsx`);
          Swal.close();
          
      } catch(e) { console.error(e); Swal.close(); alertError("Terjadi error ExcelJS: " + e.message); }
  }
  // ==============================================================
  // 2. ENGINE EXCELJS: REKAP GOLONGAN (16 KOLOM - BEBAS ERROR)
  // ==============================================================
  async function buatExcelRekapGolonganJS(res) {
      try {
          const wb = new ExcelJS.Workbook();
          const sheet = wb.addWorksheet('Rekap Golongan', { pageSetup: { paperSize: 5, orientation: 'landscape', margins: { left: 0.2, right: 0.2, top: 0.4, bottom: 0.4 } } });

          const colWidths = [5, 25, 12, 14, 14, 14, 14, 14, 16, 16, 16, 16, 16, 16, 16, 16];
          colWidths.forEach((w, i) => { sheet.getColumn(i+1).width = w; });

          sheet.mergeCells('A1:P1'); sheet.getCell('A1').value = `REKAPITULASI PENGAJUAN TPP ASN OPD ${res.setting.Nama_Dinas} PROVINSI JAMBI`;
          sheet.mergeCells('A2:P2'); sheet.getCell('A2').value = `${res.jenisASN} ${res.unitCetak}`;
          sheet.mergeCells('A3:P3'); sheet.getCell('A3').value = `Bulan : ${res.bulanBesar}`;
          for(let i=1; i<=3; i++) { 
              let c = sheet.getCell(`A${i}`); c.font = { bold: true, size: 12 }; c.alignment = { horizontal: 'center', vertical: 'middle' }; 
          }

          sheet.mergeCells('A5:A6'); sheet.getCell('A5').value = "No.";
          sheet.mergeCells('B5:B6'); sheet.getCell('B5').value = "GOLONGAN";
          sheet.mergeCells('C5:C6'); sheet.getCell('C5').value = "Jumlah\nPegawai";
          
          sheet.mergeCells('D5:H5'); sheet.getCell('D5').value = "PERHITUNGAN TPP (NETTO)";
          ['BK','PK','KK','TB','KP'].forEach((txt, i) => sheet.getCell(6, 4+i).value = txt);
          
          sheet.mergeCells('I5:I6'); sheet.getCell('I5').value = "Jumlah TPP";
          sheet.mergeCells('J5:J6'); sheet.getCell('J5').value = "BPJS 4%";
          sheet.mergeCells('K5:K6'); sheet.getCell('K5').value = "Jumlah Kotor";
          
          sheet.mergeCells('L5:O5'); sheet.getCell('L5').value = "PENGURANGAN";
          ['PPh 21','IWP 1%','BPJS 4%','Total Potongan'].forEach((txt, i) => sheet.getCell(6, 12+i).value = txt);
          
          sheet.mergeCells('P5:P6'); sheet.getCell('P5').value = "Jumlah Bersih";

          for(let i=1; i<=16; i++) { sheet.getCell(7, i).value = i.toString(); }

          for (let r = 5; r <= 7; r++) {
              let row = sheet.getRow(r);
              row.height = r === 7 ? 15 : 25;
              for(let c = 1; c <= 16; c++) {
                  let cell = row.getCell(c);
                  cell.font = { bold: r!==7, italic: r===7, size: r===7 ? 8 : 10 };
                  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                  cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: r===7 ? 'FFE9ECEF' : 'FFCFE2F3' } };
              }
          }

          let grouped = {};
          res.data.forEach(c => {
              let golDasar = c.golonganAsli.split("/")[0].trim().toUpperCase();
              let groupName = (res.jenisASN === "PPPK") ? getGroupPPPK(golDasar) : (golDasar.includes("IX") || golDasar.includes("X") ? "PPPK" : golDasar);
              if(!grouped[groupName]) grouped[groupName] = [];
              grouped[groupName].push(c);
          });

          let urutanGol = res.jenisASN === "PPPK" ? ["XVII - XIII", "XII - IX", "VIII - V", "IV - I"] : ["IV", "III", "II", "I"];
          let no = 1; let currentRow = 8; let progressCount = 0; let totalGol = urutanGol.length;
          
          // Penampung JS Grand Total
          let grandTotals = Array(14).fill(0); 

          for (let gol of urutanGol) {
              let arr = grouped[gol] || []; 
              let sum = Array(13).fill(0);
              if(arr.length > 0) {
                  for (let c of arr) {
                      let rasio = c.tppBruto > 0 ? (c.tppNettoKinerja / c.tppBruto) : 0;
                      let vals = [
                          Math.round(c.bk * rasio), Math.round(c.pk * rasio), Math.round(c.kk * rasio), Math.round(c.tb * rasio), Math.round(c.kp * rasio),
                          c.tppNettoKinerja, c.bpjs4, (c.tppNettoKinerja + c.bpjs4),
                          c.pph21TKD, c.iwp1, c.bpjs4, (c.pph21TKD + c.iwp1 + c.bpjs4), c.tppBersih
                      ];
                      vals.forEach((v, i) => { sum[i] += v; });
                  }
                  // Hitung untuk Grand Total
                  grandTotals[0] += arr.length;
                  sum.forEach((v, i) => { grandTotals[i+1] += v; });
              }
              
              let r = sheet.getRow(currentRow);
              r.height = 40;
              r.values = [no++, "GOLONGAN " + gol, arr.length, ...sum];
              r.eachCell({ includeEmpty: true }, (cell, colN) => {
                  cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                  cell.alignment = { vertical: 'middle', horizontal: colN <= 2 ? 'left' : (colN === 3 ? 'center' : 'right') };
                  if(colN > 3) cell.numFmt = '#,##0';
                  if(cell.value === undefined) cell.value = 0;
              });
              currentRow++; progressCount++;
              
              updateProgress(progressCount, totalGol); 
              await new Promise(resolve => setTimeout(resolve, 0));
          }

          let rGrand = sheet.getRow(currentRow);
          rGrand.height = 30; 
          // Masukkan angka jadi, TANPA RUMUS
          rGrand.values = ["", "Total Penghitungan TPP", ...grandTotals];
          rGrand.getCell(2).font = {bold: true};
          
          rGrand.eachCell({ includeEmpty: true }, (cell, colN) => {
              cell.font = {bold: true}; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF3F3F3' } };
              cell.border = { top: {style:'medium'}, left: {style:'thin'}, bottom: {style:'medium'}, right: {style:'thin'} };
              if(colN >= 3) cell.numFmt = '#,##0';
          });

          cetakTTD(sheet, currentRow + 3, res.setting, 16);
          unduhFile(wb, `Rekap_Golongan_${res.bulanBesar}_${res.unitCetak}.xlsx`);
          Swal.close();
          
      } catch(e) { console.error(e); Swal.close(); alertError("Terjadi error ExcelJS: " + e.message); }
  }

  // ==============================================================
  // 3. ENGINE EXCELJS: REKENING BANK (7 KOLOM)
  // ==============================================================
  async function buatExcelRekeningJS(res) {
      const wb = new ExcelJS.Workbook();
      const sheet = wb.addWorksheet('Rekening TPP', { pageSetup: { paperSize: 5 } });

      // 1. ATUR LEBAR KOLOM
      sheet.getColumn(1).width = 5;   // No
      sheet.getColumn(2).width = 40;  // Nama
      sheet.getColumn(3).width = 20;  // Rek Penerima
      sheet.getColumn(4).width = 18;  // TPP Bersih
      sheet.getColumn(5).width = 20;  // Rek Korpri
      sheet.getColumn(6).width = 15;  // Pot Korpri
      sheet.getColumn(7).width = 18;  // Jml Diterima

      // 2. MERGE CELL JUDUL & CENTER
      sheet.mergeCells('A1:G1'); sheet.getCell('A1').value = `Daftar : Pembayaran TPP BULAN ${res.bulanBesar}`;
      sheet.mergeCells('A2:G2'); sheet.getCell('A2').value = `ASN ${res.setting.Nama_Dinas} PEMERINTAH PROVINSI JAMBI`;
      sheet.mergeCells('A3:G3'); sheet.getCell('A3').value = `${res.jenisASN} ${res.unitCetak}`;
      sheet.mergeCells('A4:G4'); sheet.getCell('A4').value = `DI BANK JAMBI`;
      
      for(let i=1; i<=4; i++) { 
          sheet.getCell(`A${i}`).font = { bold: true, size: 12 }; 
          sheet.getCell(`A${i}`).alignment = { horizontal: 'center', vertical: 'middle' }; 
      }

      let rowH = sheet.getRow(6);
      rowH.height = 35;
      rowH.values = ["NO", "NAMA PEGAWAI", "NO. REK PENERIMA", "TPP BERSIH", "NO. REK KORPRI", "POT KORPRI", "JML DITERIMA"];
      rowH.eachCell({ includeEmpty: true }, cell => {
          cell.font = { bold: true }; cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'CFE2F3' } };
          cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
      });

      let no = 1; let currentRow = 7; let potKorpri = parseFloat(res.setting.Pot_Korpri) || 0;
      let progressCount = 0; let totalData = res.data.length;
      
      for (let c of res.data) {
          let golDasar = c.golonganAsli.split("/")[0].trim().toUpperCase();
          if(res.jenisASN !== "PPPK" && (golDasar.includes("IX") || golDasar.includes("X") || golDasar.includes("PPPK"))) continue;

          let jmlDiterima = c.tppBersih - potKorpri;
          let r = sheet.getRow(currentRow);
          r.height = 25;
          r.values = [ no++, c.nama, c.rekening, c.tppBersih, res.setting.Rek_Korpri, potKorpri, jmlDiterima ];
          r.eachCell({ includeEmpty: true }, (cell, colN) => {
              cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
              cell.alignment = { vertical: 'middle', horizontal: colN === 2 ? 'left' : (colN === 3 || colN === 5 ? 'center' : 'right') };
              if(colN === 4 || colN === 6 || colN === 7) cell.numFmt = '#,##0';
              if (cell.value === undefined) cell.value = 0;
          });
          currentRow++; progressCount++;
          
          updateProgress(progressCount, totalData); 
          if (progressCount % 10 === 0 || progressCount === totalData) { await new Promise(resolve => setTimeout(resolve, 0)); }
      }

      let rGrand = sheet.getRow(currentRow);
      rGrand.height = 30; rGrand.getCell('B').value = "TOTAL KESELURUHAN"; rGrand.getCell('B').font = {bold: true};
      ['D', 'F', 'G'].forEach(col => { rGrand.getCell(col).value = { formula: `SUM(${col}7:${col}${currentRow-1})` }; });
      rGrand.eachCell({ includeEmpty: true }, (cell, colN) => {
          cell.font = {bold: true}; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F3F3F3' } };
          cell.border = { top: {style:'medium'}, left: {style:'thin'}, bottom: {style:'medium'}, right: {style:'thin'} };
          if(colN === 4 || colN === 6 || colN === 7) cell.numFmt = '#,##0';
      });

      cetakTTD(sheet, currentRow + 3, res.setting, 7);
      unduhFile(wb, `Rekening_TPP_${res.bulanBesar}_${res.unitCetak}.xlsx`);
      Swal.close();
  }

  // ==============================================================
  // 4. ENGINE EXCELJS: REKAP PAJAK TER (20 KOLOM)
  // ==============================================================
  async function buatExcelRekapPajakJS(res) {
      const wb = new ExcelJS.Workbook();
      const sheet = wb.addWorksheet('Rekap Pajak TER', { pageSetup: { paperSize: 5, orientation: 'landscape' } });

      // 1. ATUR LEBAR KOLOM
      sheet.getColumn(1).width = 5;   // No
      sheet.getColumn(2).width = 22;  // NIP
      sheet.getColumn(3).width = 30;  // Nama Pegawai
      sheet.getColumn(4).width = 30;  // Unit Kerja
      sheet.getColumn(5).width = 12;  // Masa Pajak
      sheet.getColumn(6).width = 10;  // Tahun
      sheet.getColumn(7).width = 12;  // Status
      sheet.getColumn(8).width = 18;  // NPWP
      sheet.getColumn(9).width = 10;  // Kawin
      sheet.getColumn(10).width = 15; // Jenis Jabatan
      for(let i=11; i<=15; i++) sheet.getColumn(i).width = 16; // Gaji/Uang
      sheet.getColumn(16).width = 12; // Kat TER
      sheet.getColumn(17).width = 10; // Tarif
      for(let i=18; i<=20; i++) sheet.getColumn(i).width = 16; // Total Pajak

      // 2. MERGE CELL JUDUL & CENTER
      sheet.mergeCells('A1:T1'); sheet.getCell('A1').value = "REKAPITULASI PAJAK TER PENGHASILAN PEGAWAI";
      sheet.mergeCells('A2:T2'); sheet.getCell('A2').value = `PERIODE: ${res.bulanBesar}`;
      for(let i=1; i<=2; i++) { 
          sheet.getCell(`A${i}`).font = { bold: true, size: 12 }; 
          sheet.getCell(`A${i}`).alignment = { horizontal: 'center', vertical: 'middle' }; 
      }

      let rowH = sheet.getRow(4);
      rowH.height = 55;
      rowH.values = [ "No.", "NIP", "Nama Pegawai", "Unit Kerja", "Masa\nPajak", "Tahun\nPajak", "Status\nPegawai", "NPWP", "Status\nKawin", "Jenis\nJabatan", "Gaji Bruto\n(Amprah)", "PPh 21\nGaji", "Gaji Bruto\nBersih PPh", "TPP\nBruto", "Penghasilan Kotor\n(DPP TER)", "Kategori\nTER", "Tarif\n(%)", "Total PPh\nTerutang", "PPh 21\nAmprah Gaji", "Pajak TER\n(TPP)" ];
      
      rowH.eachCell({ includeEmpty: true }, cell => {
          cell.font = { bold: true }; cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FCE5CD' } }; 
          cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
      });

      res.data.sort((a, b) => {
          if (a.unitKerjaLengkap < b.unitKerjaLengkap) return -1;
          if (a.unitKerjaLengkap > b.unitKerjaLengkap) return 1;
          return b.gapokRaw - a.gapokRaw; 
      });

      let no = 1; let currentRow = 5; let progressCount = 0; let totalData = res.data.length;

      for (let c of res.data) {
          let isPPPK = String(c.golonganAsli).toUpperCase().includes("IX") || String(c.golonganAsli).toUpperCase().includes("X") || String(c.golonganAsli).toUpperCase().includes("PPPK");
          let statusPegawai = isPPPK ? "PPPK" : "PNS";
          let gajiBrutoBersihPph = c.gajiKotorTER - c.pphGajiTER;

          let r = sheet.getRow(currentRow);
          r.height = 40; 
          r.values = [ no++, c.nip, c.nama, c.unitKerjaLengkap, res.namaBulan, res.tahun, statusPegawai, "", c.statusTER, c.jabatan, c.gajiKotorTER, c.pphGajiTER, gajiBrutoBersihPph, c.tppBruto, c.dasarPajakTER, "TER " + c.katTER, c.pctTER + "%", c.pph21TotalSebulanTER, c.pphGajiTER, c.pph21TKD ];
          
          r.eachCell({ includeEmpty: true }, (cell, colN) => {
              cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
              cell.alignment = { vertical: 'middle', wrapText: true, horizontal: colN <= 4 ? 'left' : (colN <= 10 || colN === 16 || colN === 17 ? 'center' : 'right') };
              if((colN >= 11 && colN <= 15) || colN >= 18) cell.numFmt = '#,##0';
              if(cell.value === undefined) cell.value = ""; 
          });
          currentRow++; progressCount++;
          
          updateProgress(progressCount, totalData); 
          if (progressCount % 10 === 0 || progressCount === totalData) { await new Promise(resolve => setTimeout(resolve, 0)); }
      }

      let rGrand = sheet.getRow(currentRow);
      rGrand.height = 30;
      sheet.mergeCells(`B${currentRow}:D${currentRow}`); rGrand.getCell('B').value = "TOTAL KESELURUHAN"; 
      ['K','L','M','N','O','R','S','T'].forEach(col => { rGrand.getCell(col).value = { formula: `SUM(${col}5:${col}${currentRow-1})` }; });
      
      rGrand.eachCell({ includeEmpty: true }, (cell, colN) => {
          cell.font = {bold: true}; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FCE5CD' } };
          cell.border = { top: {style:'medium'}, left: {style:'thin'}, bottom: {style:'medium'}, right: {style:'thin'} };
          cell.alignment = { vertical: 'middle', horizontal: colN <= 4 ? 'center' : 'right' };
          if((colN >= 11 && colN <= 15) || colN >= 18) cell.numFmt = '#,##0';
      });

      cetakTTD(sheet, currentRow + 3, res.setting, 20);
      unduhFile(wb, `Rekap_PajakTER_${res.bulanBesar}_${res.unitCetak}.xlsx`);
      Swal.close();
  }

  // ==============================================================
  // 5. ENGINE EXCELJS: PERHITUNGAN KOLEKTIF TPP (MONSTER 63 KOLOM)
  // ==============================================================
  async function buatExcelPerhitunganJS(res) {
      const wb = new ExcelJS.Workbook();
      const sheet = wb.addWorksheet('Perhitungan TPP', { pageSetup: { paperSize: 5, orientation: 'landscape' } });

      let cols = [{ width: 5 }, { width: 35 }, { width: 8 }, { width: 8 }]; 
      for(let i=0; i<14; i++) cols.push({ width: 5 }); 
      cols.push({ width: 10 }, { width: 10 }); 
      for(let i=0; i<42; i++) cols.push({ width: 12 }); 
      cols.push({ width: 15 }); 
      sheet.columns = cols;

      sheet.mergeCells('A1:BK1'); sheet.getCell('A1').value = `PERHITUNGAN TPP ( TAMBAHAN PENGHASILAN PEGAWAI ) OPD ${res.setting.Nama_Dinas} BULAN ${res.bulanBesar}`;
      sheet.mergeCells('A2:BK2'); sheet.getCell('A2').value = `${res.jenisASN} ${res.unitCetak}`;
      for(let i=1; i<=2; i++) { sheet.getCell(`A${i}`).font = { bold: true, size: 12 }; sheet.getCell(`A${i}`).alignment = { horizontal: 'center' }; }

      sheet.mergeCells('A5:A7'); sheet.getCell('A5').value = "No.";
      sheet.mergeCells('B5:B7'); sheet.getCell('B5').value = "Nama / NIP / Gol. / Jabatan";
      sheet.mergeCells('C5:C7'); sheet.getCell('C5').value = "Kelas\nJab.";
      sheet.mergeCells('D5:D7'); sheet.getCell('D5').value = "Hari\nKerja";
      sheet.mergeCells('E5:R5'); sheet.getCell('E5').value = "Keterangan Tidak Hadir (KTH)";
      sheet.mergeCells('E6:E7'); sheet.getCell('E6').value = "DL";
      sheet.mergeCells('F6:F7'); sheet.getCell('F6').value = "S";
      sheet.mergeCells('G6:G7'); sheet.getCell('G6').value = "C";
      sheet.mergeCells('H6:H7'); sheet.getCell('H6').value = "KP";
      sheet.mergeCells('I6:L6'); sheet.getCell('I6').value = "Terlambat Datang";
      sheet.getCell('I7').value = "TL1"; sheet.getCell('J7').value = "TL2"; sheet.getCell('K7').value = "TL3"; sheet.getCell('L7').value = "TL4";
      sheet.mergeCells('M6:P6'); sheet.getCell('M6').value = "Cepat Pulang";
      sheet.getCell('M7').value = "CP1"; sheet.getCell('N7').value = "CP2"; sheet.getCell('O7').value = "CP3"; sheet.getCell('P7').value = "CP4";
      sheet.mergeCells('Q6:Q7'); sheet.getCell('Q6').value = "TK";
      sheet.mergeCells('R6:R7'); sheet.getCell('R6').value = "AS/UB";
      sheet.mergeCells('S5:S7'); sheet.getCell('S5').value = "Predikat\nKerja";
      sheet.mergeCells('T5:T7'); sheet.getCell('T5').value = "Tidak\nMenilai";

      sheet.mergeCells('U5:Z6'); sheet.getCell('U5').value = "Kriteria TPP Dasar (100%)";
      ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(20 + i) + '7').value = txt);
      sheet.mergeCells('AA5:AR5'); sheet.getCell('AA5').value = "Perhitungan Produktivitas Kerja (PK 60%)";
      sheet.mergeCells('AA6:AF6'); sheet.getCell('AA6').value = "Nilai PK (Setelah Bobot)";
      ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(26 + i) + '7').value = txt);
      sheet.mergeCells('AG6:AL6'); sheet.getCell('AG6').value = "Potongan Absen PK";
      ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(32 + i) + '7').value = txt);
      sheet.mergeCells('AM6:AR6'); sheet.getCell('AM6').value = "Akhir PK Netto";
      ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(38 + i) + '7').value = txt);
      sheet.mergeCells('AS5:BJ5'); sheet.getCell('AS5').value = "Perhitungan Disiplin Kerja (DK 40%)";
      sheet.mergeCells('AS6:AX6'); sheet.getCell('AS6').value = "Bobot DK Dasar (40%)";
      ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(44 + i) + '7').value = txt);
      sheet.mergeCells('AY6:BD6'); sheet.getCell('AY6').value = "Potongan Absen DK";
      ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(50 + i) + '7').value = txt);
      sheet.mergeCells('BE6:BJ6'); sheet.getCell('BE6').value = "Akhir DK Netto";
      ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(56 + i) + '7').value = txt);
      sheet.mergeCells('BK5:BK7'); sheet.getCell('BK5').value = "TPP Diterima\nSebelum Pajak";

      for(let i=1; i<=63; i++) { sheet.getCell(numToLet(i-1) + '8').value = i.toString(); }

      for (let i = 5; i <= 8; i++) {
          sheet.getRow(i).eachCell({ includeEmpty: true }, (cell, colN) => {
              cell.font = { bold: i!==8, italic: i===8, size: i===8 ? 8 : 9 }; 
              cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
              cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
              
              if(i===8) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E9ECEF' } };
              else if(colN >= 5 && colN <= 18) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9EAD3' } }; 
              else if(colN >= 27 && colN <= 44) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FCE5CD' } }; 
              else if(colN >= 45 && colN <= 62) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E4D7F5' } }; 
              else cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'CFE2F3' } }; 
          });
      }
      sheet.getRow(5).height = 25; sheet.getRow(6).height = 25; sheet.getRow(7).height = 30; sheet.getRow(8).height = 15;

      let grouped = {};
      res.data.forEach(c => {
          let golDasar = c.golonganAsli.split("/")[0].trim().toUpperCase();
          let groupName = (res.jenisASN === "PPPK") ? getGroupPPPK(golDasar) : (golDasar.includes("IX") || golDasar.includes("X") ? "PPPK" : golDasar);
          if(!grouped[groupName]) grouped[groupName] = [];
          grouped[groupName].push(c);
      });

      let urutanGol = res.jenisASN === "PPPK" ? ["XVII - XIII", "XII - IX", "VIII - V", "IV - I"] : ["IV", "III", "II", "I"];
      let noUrut = 1; let currentRow = 9; let subTotalRows = [];
      let progressCount = 0; let totalData = res.data.length;
      let deteksiKolomNol = Array(64).fill(0);

      for (let gol of urutanGol) {
          if(grouped[gol] && grouped[gol].length > 0) {
              let startRowGroup = currentRow;
              for (let c of grouped[gol]) {
                  let a = c.rawAbsen;
                  let kr = [c.bk, c.pk, c.kk, c.tb, c.kp, c.tppBruto];
                  let nSKP = kr.map(v => v * 0.6 * c.rasioSKP);
                  let pAbsPK = nSKP.map(v => v * c.pctAbsPK);
                  let akhPK = nSKP.map((v, i) => v - pAbsPK[i]);
                  let bDK = kr.map(v => v * 0.4);
                  let pAbsDK = bDK.map(v => v * c.pctAbsDK);
                  let akhDK = bDK.map((v, i) => v - pAbsDK[i]);

                  let rValues = [
                      noUrut++, `${c.nama}\nNIP. ${c.nip}\n${res.jenisASN} - Gol. ${c.golonganAsli}\n${c.jabatan}`, c.pergub ? c.pergub.kelasJabatan : "-", c.hk,
                      a.dl, a.s, a.c, a.kp, a.tl1, a.tl2, a.tl3, a.tl4, a.cp1, a.cp2, a.cp3, a.cp4, a.tk, a.asub,
                      c.skp, c.menilai === "Tidak Menilai" ? "Tidak Menilai" : "",
                      ...kr, ...nSKP, ...pAbsPK, ...akhPK, ...bDK, ...pAbsDK, ...akhDK,
                      (akhPK[5] + akhDK[5])
                  ];

                  let row = sheet.getRow(currentRow);
                  row.height = 85; 
                  row.values = rValues;
                  
                  for(let k = 20; k < 63; k++) { deteksiKolomNol[k + 1] += parseFloat(rValues[k]) || 0; }

                  row.eachCell({ includeEmpty: true }, (cell, colN) => {
                      cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                      cell.alignment = { vertical: 'middle', wrapText: true, horizontal: colN === 2 ? 'left' : (colN <= 20 ? 'center' : 'right') };
                      if(colN >= 21) cell.numFmt = '#,##0';
                      if(cell.value === undefined && colN >= 21) cell.value = 0;
                  });
                  currentRow++; progressCount++;
                  
                  updateProgress(progressCount, totalData); 
                  if (progressCount % 5 === 0 || progressCount === totalData) { await new Promise(r => setTimeout(r, 10)); }
              }

              let rowSub = sheet.getRow(currentRow);
              rowSub.height = 25; rowSub.getCell('B').value = `TOTAL GOLONGAN ${gol}`; rowSub.getCell('B').font = { bold: true };
              for(let i=21; i<=63; i++) { rowSub.getCell(i).value = { formula: `SUM(${numToLet(i-1)}${startRowGroup}:${numToLet(i-1)}${currentRow-1})` }; }

              rowSub.eachCell({ includeEmpty: true }, (cell, colN) => {
                  cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F3F3F3' } };
                  if (colN >= 21) { cell.numFmt = '#,##0'; cell.font = {bold: true}; }
              });
              subTotalRows.push(currentRow); currentRow++;
          }
      }

      let rowGrand = sheet.getRow(currentRow);
      rowGrand.height = 30; rowGrand.getCell('B').value = "TOTAL KESELURUHAN"; rowGrand.getCell('B').font = { bold: true };
      if(subTotalRows.length > 0) { for(let i=21; i<=63; i++) { rowGrand.getCell(i).value = { formula: subTotalRows.map(rn => `${numToLet(i-1)}${rn}`).join('+') }; } }
      
      rowGrand.eachCell({ includeEmpty: true }, (cell, colN) => {
          cell.border = { top: {style:'medium'}, left: {style:'thin'}, bottom: {style:'medium'}, right: {style:'thin'} };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E9ECEF' } };
          if (colN >= 21) { cell.numFmt = '#,##0'; cell.font = {bold: true}; }
      });

      sheet.getColumn(1).width = 5; sheet.getColumn(2).width = 35; sheet.getColumn(3).width = 8; sheet.getColumn(4).width = 8;
      for(let i=5; i<=18; i++) sheet.getColumn(i).width = 5; 
      sheet.getColumn(19).width = 10; sheet.getColumn(20).width = 10; 

      for(let i=21; i<=63; i++) {
          if (deteksiKolomNol[i] === 0) sheet.getColumn(i).width = 3.5; 
          else sheet.getColumn(i).width = 12;  
      }
      sheet.getColumn(63).width = 15;

      cetakTTD(sheet, currentRow + 3, res.setting, 63); 
      unduhFile(wb, `Perhitungan_TPP_${res.bulanBesar}_${res.unitCetak}.xlsx`);
      Swal.close();
  }

