function showDownloadModal(url) { let modalEl = document.getElementById('modalDownloadReady'); if(modalEl) { document.getElementById('btnRealDownload').href = url; let modalObj = bootstrap.Modal.getInstance(modalEl) || new bootstrap.Modal(modalEl); modalObj.show(); } else { window.open(url, '_blank'); } }
function tutupModalDownload() { setTimeout(() => { let modalEl = document.getElementById('modalDownloadReady'); let modal = bootstrap.Modal.getInstance(modalEl); if(modal) modal.hide(); }, 1000); }

function cetakPDFPerorangan(tabId, namaFile) { window.scrollTo(0, 0); let element = document.getElementById('viewManajemenASN'); let navPills = element.querySelector('.nav-pills'); let actionBtns = element.querySelectorAll('button'); let origPillsDisplay = navPills ? navPills.style.display : ''; if(navPills) navPills.style.display = 'none'; actionBtns.forEach(b => { b.dataset.origDisplay = b.style.display; b.style.display = 'none'; }); let tables = element.querySelectorAll('.table-responsive'); tables.forEach(t => { t.style.overflow = 'visible'; }); let selects = element.querySelectorAll('.form-select'); selects.forEach(s => { s.dataset.origBg = s.style.backgroundImage; s.style.backgroundImage = 'none'; }); startLoading("Menyusun PDF..."); let opt = { margin: 0.3, filename: `${namaFile}_${asNIPAktif}.pdf`, image: { type: 'jpeg', quality: 0.98 }, html2canvas: { scale: 2, useCORS: true, scrollY: 0, scrollX: 0 }, jsPDF: { unit: 'in', format: 'legal', orientation: 'portrait' } }; html2pdf().set(opt).from(element).save().then(() => { stopLoading(); if(navPills) navPills.style.display = origPillsDisplay; actionBtns.forEach(b => { b.style.display = b.dataset.origDisplay; }); tables.forEach(t => { t.style.overflow = ''; }); selects.forEach(s => { s.style.backgroundImage = s.dataset.origBg; }); setTimeout(() => { alertSukses("File PDF Perorangan berhasil diunduh!"); }, 500); }).catch(err => { stopLoading(); if(navPills) navPills.style.display = origPillsDisplay; actionBtns.forEach(b => { b.style.display = b.dataset.origDisplay; }); tables.forEach(t => { t.style.overflow = ''; }); selects.forEach(s => { s.style.backgroundImage = s.dataset.origBg; }); setTimeout(() => { alertError("Gagal memproses PDF: " + err); }, 500); }); }

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
        statusEl.innerText = "2. Menyusun File...";
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
// TRIGGER DOWNLOAD
// ==============================================================
async function fetchDataLaporan(format, e, funcExcel, funcPdf) {
    if(e) e.preventDefault(); 
    if(!globalBulanAktif) return alertPeringatan("Pilih bulan terlebih dahulu!"); 
    
    let fUnit = document.getElementById('filterUnitKerja').value; 
    
    showLoadingPhase1(); 
    try {
        let actionName = (funcExcel === buatExcelNominatifJS) ? "getJsonNominatif" : "getJsonLaporanLengkap";
        let payload = { bulanAktif: globalBulanAktif, jenisASN: globalJenisASN, roleUser: currentUser.role, unitkerjaUser: currentUser.unitkerja, filterUnit: fUnit };
        
        let res = await fetchAPI(actionName, payload); 
        
        if(res && res.error) {
            Swal.close(); return alertError("❌ " + res.error); 
        } else if (res && res.status === "sukses") {
            await new Promise(r => setTimeout(r, 100));
            let adaBelumSimpan = res.data.some(c => (c.gajiKotor === 0 || c.gajiKotorTER === 0));
            
            if (adaBelumSimpan) {
                Swal.close(); 
                Swal.fire({
                    title: 'Ada Data Belum Lengkap!',
                    text: 'Ditemukan pegawai yang data Gaji-nya belum tersimpan (angkanya Rp 0). Yakin ingin lanjut mencetak?',
                    icon: 'warning', showCancelButton: true, confirmButtonColor: '#3085d6', cancelButtonColor: '#d33',
                    confirmButtonText: 'Ya, Lanjut Cetak!', cancelButtonText: 'Batal'
                }).then(async (result) => {
                    if (result.isConfirmed) {
                        showLoadingPhase1(); 
                        if(format === 'excel') { 
                            switchToPhase2(res.data.length); 
                            await new Promise(r => setTimeout(r, 100)); 
                            await funcExcel(res); 
                        } else { funcPdf(res); } 
                    }
                });
            } else {
                if(format === 'excel') { 
                    switchToPhase2(res.data.length); 
                    await new Promise(r => setTimeout(r, 100)); 
                    await funcExcel(res); 
                } else { funcPdf(res); } 
            }
        } else { Swal.close(); alertError("❌ Gagal mengambil data dari server. Periksa koneksi."); } 
    } catch (err) { Swal.close(); alertError("Terjadi kesalahan sistem: " + err.message); }
}

function unduhNominatifKolektif(format, e) { fetchDataLaporan(format, e, buatExcelNominatifJS, buatPdfNominatifJS); } 
function unduhPerhitunganKolektif(format, e) { fetchDataLaporan(format, e, buatExcelPerhitunganJS, buatPdfPerhitunganJS); }
function unduhRekapGolongan(format, e) { fetchDataLaporan(format, e, buatExcelRekapGolonganJS, buatPdfRekapGolonganJS); }
function unduhRekening(format, e) { fetchDataLaporan(format, e, buatExcelRekeningJS, buatPdfRekeningJS); }
function unduhRekapPajak(format, e) { fetchDataLaporan(format, e, buatExcelRekapPajakJS, buatPdfRekapPajakJS); }

const getNum = (val) => {
    let num = parseFloat(val);
    return (isNaN(num) || !isFinite(num)) ? 0 : num;
};

const formatDigit = (num) => Math.round(num).toLocaleString('id-ID'); 

// ==========================================
// HELPER TTD AMAN (POSISI DIGESER + MULTI UNIT)
// ==========================================
function cetakTTDAman(sheet, startRow, setting, colMax, namaUnitCetak) {
    const getStr = (val) => (val === null || val === undefined) ? "-" : String(val).trim();
    
    // Mesin Pencari TTD Spesifik Unit
    const getVal = (key) => {
        if (namaUnitCetak && setting[`${namaUnitCetak}_${key}`]) return setting[`${namaUnitCetak}_${key}`];
        return setting[key];
    };

    const letKiri = 'C'; 
    let idxKanan = parseInt(colMax) - 5;
    if (idxKanan <= 2) { idxKanan = parseInt(colMax) - 2; } 
    const letKanan = numToLet(idxKanan);
    
    sheet.getCell(`${letKanan}${startRow}`).value = `Jambi, ........................ ${new Date().getFullYear()}`;
    sheet.getCell(`${letKiri}${startRow+1}`).value = "Mengetahui,";
    
    sheet.getCell(`${letKiri}${startRow+2}`).value = getStr(getVal('Kepala_Jabatan'));
    sheet.getCell(`${letKanan}${startRow+2}`).value = "BENDAHARA PENGELUARAN";
    
    sheet.getCell(`${letKiri}${startRow+6}`).value = getStr(getVal('Kepala_Nama'));
    sheet.getCell(`${letKiri}${startRow+6}`).font = { bold: true, underline: true };
    sheet.getCell(`${letKanan}${startRow+6}`).value = getStr(getVal('Bendahara_Nama'));
    sheet.getCell(`${letKanan}${startRow+6}`).font = { bold: true, underline: true };
    
    sheet.getCell(`${letKiri}${startRow+7}`).value = getStr(getVal('Kepala_Pangkat'));
    sheet.getCell(`${letKanan}${startRow+7}`).value = getStr(getVal('Bendahara_Pangkat'));
    
    sheet.getCell(`${letKiri}${startRow+8}`).value = "NIP. " + getStr(getVal('Kepala_NIP'));
    sheet.getCell(`${letKanan}${startRow+8}`).value = "NIP. " + getStr(getVal('Bendahara_NIP'));

    for(let i=0; i<=8; i++) {
        sheet.getCell(`${letKiri}${startRow+i}`).alignment = { horizontal: 'left' };
        sheet.getCell(`${letKanan}${startRow+i}`).alignment = { horizontal: 'left' };
    }
}

// ==============================================================
// 1. EXCEL NOMINATIF
// ==============================================================
async function buatExcelNominatifJS(res) {
    try {
        const wb = new ExcelJS.Workbook();
        const sheet = wb.addWorksheet('Daftar Nominatif', { pageSetup: { paperSize: 5, orientation: 'landscape' } });
        let cols = [ { width: 5 }, { width: 35 }, { width: 15 }, { width: 30 }, { width: 15 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 15 }, { width: 12 }, { width: 15 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 15 }, { width: 18 } ];
        sheet.columns = cols;

        sheet.mergeCells('A1:R1'); sheet.getCell('A1').value = `DAFTAR NOMINATIF TPP ${res.setting.Nama_Dinas || ''} PEMERINTAH PROVINSI JAMBI`;
        sheet.mergeCells('A2:R2'); sheet.getCell('A2').value = `${res.jenisASN || ''} ${res.unitCetak || ''}`;
        sheet.mergeCells('A3:R3'); sheet.getCell('A3').value = `PERIODE BULAN: ${res.bulanBesar || ''}`;
        for(let i=1; i<=3; i++) { sheet.getCell(`A${i}`).font = { bold: true, size: i===1 ? 14 : 12 }; sheet.getCell(`A${i}`).alignment = { horizontal: 'center' }; }

        sheet.mergeCells('A5:A6'); sheet.getCell('A5').value = "No.";
        sheet.mergeCells('B5:B6'); sheet.getCell('B5').value = "Nama / Tgl Lahir / NIP / Gol.";
        sheet.mergeCells('C5:C6'); sheet.getCell('C5').value = "Status / Jiwa";
        sheet.mergeCells('D5:D6'); sheet.getCell('D5').value = "Jabatan";
        sheet.mergeCells('E5:E6'); sheet.getCell('E5').value = "Gaji Kotor";
        sheet.mergeCells('F5:M5'); sheet.getCell('F5').value = "PERHITUNGAN TPP (NETTO)";
        sheet.getCell('F6').value = 'BK'; sheet.getCell('G6').value = 'PK'; sheet.getCell('H6').value = 'KK'; sheet.getCell('I6').value = 'TB'; sheet.getCell('J6').value = 'KP'; sheet.getCell('K6').value = 'Jml TPP'; sheet.getCell('L6').value = 'BPJS 4%'; sheet.getCell('M6').value = 'Total Netto';
        sheet.mergeCells('N5:Q5'); sheet.getCell('N5').value = "PENGURANGAN TPP";
        sheet.getCell('N6').value = 'IWP 1%'; sheet.getCell('O6').value = 'PPh 21'; sheet.getCell('P6').value = 'BPJS 4%'; sheet.getCell('Q6').value = 'TOTAL POT.';
        sheet.mergeCells('R5:R6'); sheet.getCell('R5').value = "BERSIH DITERIMA";

        for (let i = 5; i <= 6; i++) {
            sheet.getRow(i).eachCell({ includeEmpty: true }, (cell) => { cell.font = { bold: true, size: 10 }; cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'CFE2F3' } }; });
        }
        sheet.getRow(5).height = 25; sheet.getRow(6).height = 25;

        let groups = res.jenisASN === "PPPK" ? ["XVII - XIII", "XII - IX", "VIII - V", "IV - I"] : ["IV", "III", "II", "I"];
        let rekapData = {}; groups.forEach(g => rekapData[g] = []); 

        res.data.forEach(c => {
            let golDasar = String(c.golonganAsli || "").split("/")[0].trim().toUpperCase();
            let groupName = (res.jenisASN === "PPPK") ? getGroupPPPK(golDasar) : (golDasar.includes("IX") || golDasar.includes("X") ? "PPPK" : golDasar);
            if (rekapData[groupName]) rekapData[groupName].push(c);
        });

        let curRow = 7; let no = 1; let grandTotals = Array(14).fill(0); let maxChars = Array(19).fill(0); let subTotalRows = [];

        groups.forEach(g => {
            let arr = rekapData[g];
            if (arr && arr.length > 0) {
                let subTotals = Array(14).fill(0); let startRowGroup = curRow;

                arr.forEach(c => {
                    let isKawin = String(c.statusTER || "").startsWith("K");
                    let jmlJiwa = (isKawin ? 2 : 1) + (parseInt(c.tanggungAnak) || 0);
                    let tppBruto = parseFloat(c.tppBruto) || 0; let tppNetto = parseFloat(c.tppNettoKinerja) || 0; let rasio = tppBruto > 0 ? (tppNetto / tppBruto) : 0;
                    let vGaji = parseFloat(c.gajiKotor) || 0; let vBk = Math.round((parseFloat(c.bk) || 0) * rasio); let vPk = Math.round((parseFloat(c.pk) || 0) * rasio); let vKk = Math.round((parseFloat(c.kk) || 0) * rasio); let vTb = Math.round((parseFloat(c.tb) || 0) * rasio); let vKp = Math.round((parseFloat(c.kp) || 0) * rasio); let vBpjs = parseFloat(c.bpjs4) || 0; let vIwp1 = parseFloat(c.iwp1) || 0; let vPph21 = parseFloat(c.pph21TKD) || 0;
                    let rawTpp = vBk + vPk + vKk + vTb + vKp; let rawNetto = rawTpp + vBpjs; let rawPot = vIwp1 + vPph21 + vBpjs; let rawBersih = rawNetto - rawPot;

                    let rValues = [ no++, `${c.nama || "-"}\n${c.tglLahir || "-"}\nNIP. ${c.nip || "-"}\n${res.jenisASN || "-"} - Gol. ${c.golonganAsli || "-"}`, `${c.statusTER || "-"}\nJiwa: ${jmlJiwa}`, c.jabatan || "-", vGaji, vBk, vPk, vKk, vTb, vKp, { formula: `SUM(F${curRow}:J${curRow})` }, vBpjs, { formula: `K${curRow}+L${curRow}` }, vIwp1, vPph21, vBpjs, { formula: `SUM(N${curRow}:P${curRow})` }, { formula: `M${curRow}-Q${curRow}` } ];
                    let rawDataForLength = [0, 0, 0, 0, vGaji, vBk, vPk, vKk, vTb, vKp, rawTpp, vBpjs, rawNetto, vIwp1, vPph21, vBpjs, rawPot, rawBersih];
                    let row = sheet.getRow(curRow); row.height = 85; row.values = rValues;

                    row.eachCell({ includeEmpty: true }, (cell, colN) => {
                        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                        cell.alignment = { vertical: 'middle', wrapText: true, horizontal: (colN===2||colN===4) ? 'left' : (colN>=5 ? 'right' : 'center') };
                        if (colN >= 5) { cell.numFmt = '#,##0'; subTotals[colN - 5] += rawDataForLength[colN - 1]; grandTotals[colN - 5] += rawDataForLength[colN - 1]; let len = Math.round(rawDataForLength[colN - 1]).toLocaleString('id-ID').length; maxChars[colN] = Math.max(maxChars[colN], len); }
                    });
                    curRow++;
                });

                let rowSub = sheet.getRow(curRow); rowSub.height = 25; sheet.mergeCells(`A${curRow}:D${curRow}`); rowSub.getCell(1).value = `SUB-TOTAL GOLONGAN ${g}`;
                for(let c = 5; c <= 18; c++) { rowSub.getCell(c).value = { formula: `SUM(${numToLet(c - 1)}${startRowGroup}:${numToLet(c - 1)}${curRow-1})` }; }
                rowSub.eachCell({ includeEmpty: true }, (cell, colN) => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F3F3F3' } }; cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; if (colN >= 5 || colN === 1) { cell.font = { bold: true }; cell.alignment = { horizontal: colN === 1 ? 'center' : 'right', vertical: 'middle' }; if(colN >= 5) cell.numFmt = '#,##0'; } });
                subTotalRows.push(curRow); curRow++;
            }
        });

        let rowGrand = sheet.getRow(curRow); rowGrand.height = 30; sheet.mergeCells(`A${curRow}:D${curRow}`); rowGrand.getCell(1).value = "TOTAL KESELURUHAN";
        for(let c = 5; c <= 18; c++) { let formulaStr = subTotalRows.length > 0 ? subTotalRows.map(r => `${numToLet(c - 1)}${r}`).join('+') : "0"; rowGrand.getCell(c).value = { formula: formulaStr }; }
        rowGrand.eachCell({ includeEmpty: true }, (cell, colN) => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E9ECEF' } }; cell.border = { top: {style:'medium'}, left: {style:'thin'}, bottom: {style:'medium'}, right: {style:'thin'} }; if (colN >= 5 || colN === 1) { cell.font = { bold: true }; cell.alignment = { horizontal: colN === 1 ? 'center' : 'right', vertical: 'middle' }; if(colN >= 5) cell.numFmt = '#,##0'; } });

        for (let c = 5; c <= 18; c++) { if (grandTotals[c - 5] === 0) { sheet.getColumn(c).width = 4; } else { sheet.getColumn(c).width = Math.max(10, (maxChars[c] * 1.15) + 1.5); } }

        cetakTTDAman(sheet, curRow + 3, res.setting, 18, res.unitCetak);
        unduhFile(wb, `Nominatif_TPP_${res.bulanBesar}_${res.unitCetak}.xlsx`);
    } catch (error) { Swal.close(); alertError("Gagal merakit Excel: " + error.message); }
}

// ==============================================================
// 2. EXCEL REKAP GOLONGAN (18 KOLOM - SAMA PERSIS DGN NOMINATIF)
// ==============================================================
async function buatExcelRekapGolonganJS(res) {
    try {
        const wb = new ExcelJS.Workbook();
        const sheet = wb.addWorksheet('Rekap Golongan', { pageSetup: { paperSize: 5, orientation: 'landscape' } });

        // Total 18 Kolom (Sama persis ukurannya dgn Nominatif)
        let cols = [ 
            { width: 5 }, { width: 25 }, { width: 15 }, { width: 5 }, { width: 15 }, 
            { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, 
            { width: 15 }, { width: 12 }, { width: 15 }, { width: 12 }, { width: 12 }, 
            { width: 12 }, { width: 15 }, { width: 18 } 
        ];
        sheet.columns = cols;

        sheet.mergeCells('A1:R1'); sheet.getCell('A1').value = `REKAPITULASI PENGAJUAN TPP ASN OPD ${res.setting.Nama_Dinas || ''}`;
        sheet.mergeCells('A2:R2'); sheet.getCell('A2').value = `${res.jenisASN || ''} ${res.unitCetak || ''}`;
        sheet.mergeCells('A3:R3'); sheet.getCell('A3').value = `Bulan : ${res.bulanBesar || ''}`;
        for (let i = 1; i <= 3; i++) { 
            sheet.getCell(`A${i}`).font = { bold: true, size: i === 1 ? 14 : 12 }; 
            sheet.getCell(`A${i}`).alignment = { horizontal: 'center' }; 
        }

        // --- HEADER SAMA PERSIS DENGAN NOMINATIF ---
        sheet.mergeCells('A5:A6'); sheet.getCell('A5').value = "No.";
        sheet.mergeCells('B5:B6'); sheet.getCell('B5').value = "GOLONGAN";
        sheet.mergeCells('C5:D6'); sheet.getCell('C5').value = "Jumlah Pegawai";
        sheet.mergeCells('E5:E6'); sheet.getCell('E5').value = "Gaji Kotor";
        
        sheet.mergeCells('F5:M5'); sheet.getCell('F5').value = "PERHITUNGAN TPP (NETTO)";
        sheet.getCell('F6').value = 'BK'; sheet.getCell('G6').value = 'PK'; sheet.getCell('H6').value = 'KK';
        sheet.getCell('I6').value = 'TB'; sheet.getCell('J6').value = 'KP'; sheet.getCell('K6').value = 'Jml TPP';
        sheet.getCell('L6').value = 'BPJS 4%'; sheet.getCell('M6').value = 'Total Netto';

        sheet.mergeCells('N5:Q5'); sheet.getCell('N5').value = "PENGURANGAN TPP";
        sheet.getCell('N6').value = 'IWP 1%'; sheet.getCell('O6').value = 'PPh 21';
        sheet.getCell('P6').value = 'BPJS 4%'; sheet.getCell('Q6').value = 'TOTAL POT.';

        sheet.mergeCells('R5:R6'); sheet.getCell('R5').value = "BERSIH DITERIMA";

        for (let i = 5; i <= 6; i++) {
            sheet.getRow(i).eachCell({ includeEmpty: true }, (cell) => { 
                cell.font = { bold: true, size: 10 }; 
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; 
                cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; 
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'CFE2F3' } }; 
            });
        }
        sheet.getRow(5).height = 25; sheet.getRow(6).height = 25;

        // --- PROSES PENGELOMPOKAN ---
        let groups = res.jenisASN === "PPPK" ? ["XVII - XIII", "XII - IX", "VIII - V", "IV - I"] : ["IV", "III", "II", "I"];
        let rekap = {}; 
        groups.forEach(g => { rekap[g] = { peg:0, gaji:0, bk:0, pk:0, kk:0, tb:0, kp:0, bpjs:0, iwp:0, pph:0, bersih:0 }; });

        res.data.forEach(c => {
            let golDasar = String(c.golonganAsli || "").split("/")[0].trim().toUpperCase();
            let groupName = (res.jenisASN === "PPPK") ? getGroupPPPK(golDasar) : (golDasar.includes("IX") || golDasar.includes("X") ? "PPPK" : golDasar);
            if(!rekap[groupName]) rekap[groupName] = { peg:0, gaji:0, bk:0, pk:0, kk:0, tb:0, kp:0, bpjs:0, iwp:0, pph:0, bersih:0 };
            
            let tppBruto = parseFloat(c.tppBruto) || 0; 
            let tppNetto = parseFloat(c.tppNettoKinerja) || 0; 
            let rasio = tppBruto > 0 ? (tppNetto / tppBruto) : 0;
            
            rekap[groupName].peg++; 
            rekap[groupName].gaji += parseFloat(c.gajiKotor) || 0;
            rekap[groupName].bk += Math.round((parseFloat(c.bk) || 0) * rasio); 
            rekap[groupName].pk += Math.round((parseFloat(c.pk) || 0) * rasio); 
            rekap[groupName].kk += Math.round((parseFloat(c.kk) || 0) * rasio); 
            rekap[groupName].tb += Math.round((parseFloat(c.tb) || 0) * rasio); 
            rekap[groupName].kp += Math.round((parseFloat(c.kp) || 0) * rasio); 
            rekap[groupName].bpjs += parseFloat(c.bpjs4) || 0;
            rekap[groupName].iwp += parseFloat(c.iwp1) || 0;
            rekap[groupName].pph += parseFloat(c.pph21TKD) || 0;
            rekap[groupName].bersih += parseFloat(c.tppBersih) || 0;
        });

        let curRow = 7; let no = 1; let grandTotals = Array(18).fill(0); let maxChars = Array(19).fill(0);

        // --- CETAK 4 BARIS GOLONGAN ---
        groups.forEach(g => {
            let dat = rekap[g];
            
            // Variabel Mentah Untuk Lebar Kolom
            let rawTpp = dat.bk + dat.pk + dat.kk + dat.tb + dat.kp;
            let rawNetto = rawTpp + dat.bpjs;
            let rawPot = dat.iwp + dat.pph + dat.bpjs;
            let rawBersih = rawNetto - rawPot;

            // RUMUS 100% SAMA DENGAN NOMINATIF
            let rValues = [ 
                no++, `GOLONGAN ${g}`, dat.peg, "", dat.gaji, 
                dat.bk, dat.pk, dat.kk, dat.tb, dat.kp, 
                { formula: `SUM(F${curRow}:J${curRow})` }, // K: Jml TPP
                dat.bpjs, 
                { formula: `K${curRow}+L${curRow}` },     // M: Total Netto
                dat.iwp, dat.pph, dat.bpjs, 
                { formula: `SUM(N${curRow}:P${curRow})` }, // Q: Total Pot
                { formula: `M${curRow}-Q${curRow}` }       // R: Bersih Diterima
            ];
            
            let rawDataForLength = [0, 0, dat.peg, 0, dat.gaji, dat.bk, dat.pk, dat.kk, dat.tb, dat.kp, rawTpp, dat.bpjs, rawNetto, dat.iwp, dat.pph, dat.bpjs, rawPot, rawBersih];
            
            let row = sheet.getRow(curRow); row.height = 30; row.values = rValues;

            // Merge Kolom C dan D biar teks "Jumlah Pegawai" ditengah
            sheet.mergeCells(`C${curRow}:D${curRow}`);

            row.eachCell({ includeEmpty: true }, (cell, colN) => {
                cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; 
                cell.alignment = { vertical: 'middle', horizontal: colN <= 2 ? 'left' : (colN === 3 || colN === 4 ? 'center' : 'right') };
                if (colN >= 5 || colN === 3) { 
                    cell.numFmt = '#,##0'; 
                    grandTotals[colN - 1] += rawDataForLength[colN - 1]; 
                    let len = Math.round(rawDataForLength[colN - 1]).toLocaleString('id-ID').length; 
                    maxChars[colN] = Math.max(maxChars[colN], len); 
                }
            });
            curRow++;
        });

        // --- CETAK GRAND TOTAL ---
        let rowGrand = sheet.getRow(curRow); rowGrand.height = 30; 
        sheet.mergeCells(`A${curRow}:B${curRow}`); rowGrand.getCell(1).value = "TOTAL KESELURUHAN"; 
        sheet.mergeCells(`C${curRow}:D${curRow}`); rowGrand.getCell(3).value = { formula: `SUM(C7:C${curRow-1})` };
        
        for(let c = 5; c <= 18; c++) { 
            rowGrand.getCell(c).value = { formula: `SUM(${numToLet(c - 1)}7:${numToLet(c - 1)}${curRow-1})` }; 
        }
        
        rowGrand.eachCell({ includeEmpty: true }, (cell, colN) => { 
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E9ECEF' } }; 
            cell.border = { top: {style:'medium'}, left: {style:'thin'}, bottom: {style:'medium'}, right: {style:'thin'} }; 
            if (colN >= 3 || colN === 1) { 
                cell.font = { bold: true }; 
                cell.alignment = { horizontal: colN === 1 ? 'center' : (colN === 3 || colN === 4 ? 'center' : 'right'), vertical: 'middle' }; 
                if (colN >= 5) cell.numFmt = '#,##0'; 
            } 
        });

        // --- AUTOFIT & SHRINK KOLOM KOSONG ---
        for (let c = 5; c <= 18; c++) { 
            if (grandTotals[c - 1] === 0) { 
                sheet.getColumn(c).width = 4; // Ciutkan jika 0 Rupiah
            } else { 
                sheet.getColumn(c).width = Math.max(10, (maxChars[c] * 1.15) + 1.5); 
            } 
        }

        // --- POSISI TTD AMAN (SAMA SEPERTI NOMINATIF, 18 KOLOM) ---
        cetakTTDAman(sheet, curRow + 3, res.setting, 18, res.unitCetak);
        
        unduhFile(wb, `Rekap_Golongan_${res.bulanBesar}_${res.unitCetak}.xlsx`);
    } catch (error) { 
        Swal.close(); 
        alertError("Gagal merakit Excel Rekap Golongan: " + error.message); 
    }
}


// ==============================================================
// 3. EXCEL REKENING (FIX RUMUS TPP BERSIH SAMA DGN NOMINATIF)
// ==============================================================
async function buatExcelRekeningJS(res) {
    try {
        const wb = new ExcelJS.Workbook();
        const sheet = wb.addWorksheet('Rekening TPP', { pageSetup: { paperSize: 5 } });
        sheet.columns = [ {width: 5}, {width: 40}, {width: 20}, {width: 18}, {width: 20}, {width: 15}, {width: 18} ];

        sheet.mergeCells('A1:G1'); sheet.getCell('A1').value = `Daftar : Pembayaran TPP BULAN ${res.bulanBesar}`;
        sheet.mergeCells('A2:G2'); sheet.getCell('A2').value = `ASN ${res.setting.Nama_Dinas} PEMERINTAH PROVINSI JAMBI`;
        sheet.mergeCells('A3:G3'); sheet.getCell('A3').value = `${res.jenisASN} ${res.unitCetak}`;
        sheet.mergeCells('A4:G4'); sheet.getCell('A4').value = `DI BANK JAMBI`;
        for(let i=1; i<=4; i++) { 
            sheet.getCell(`A${i}`).font = { bold: true, size: 12 }; 
            sheet.getCell(`A${i}`).alignment = { horizontal: 'center', vertical: 'middle' }; 
        }

        let rowH = sheet.getRow(6); rowH.height = 35; 
        rowH.values = ["NO", "NAMA PEGAWAI", "NO. REK PENERIMA", "TPP BERSIH", "NO. REK KORPRI", "POT KORPRI", "JML DITERIMA"];
        rowH.eachCell({ includeEmpty: true }, cell => { 
            cell.font = { bold: true }; 
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; 
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'CFE2F3' } }; 
            cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; 
        });

        let no = 1; let currentRow = 7; 
        let potKorpri = parseFloat(res.setting.Pot_Korpri) || 0; 
        let progressCount = 0; let totalData = res.data.length;
        
        for (let c of res.data) {
            let golDasar = String(c.golonganAsli).split("/")[0].trim().toUpperCase();
            if(res.jenisASN !== "PPPK" && (golDasar.includes("IX") || golDasar.includes("X") || golDasar.includes("PPPK"))) continue;

            // 👇 INI KUNCINYA: Rumus TPP Bersih Disamakan Persis dengan Kolom (R) Nominatif
            let tppNettoKinerja = parseFloat(c.tppNettoKinerja) || 0;
            let iwp1 = parseFloat(c.iwp1) || 0;
            let pph21TKD = parseFloat(c.pph21TKD) || 0;
            
            // Rumus Nominatif: TPP Netto - IWP 1% - PPh 21
            let tppBersihNominatif = tppNettoKinerja - iwp1 - pph21TKD;
            
            let jmlDiterima = tppBersihNominatif - potKorpri; 
            
            let r = sheet.getRow(currentRow); r.height = 25;
            r.values = [ no++, c.nama, c.rekening, tppBersihNominatif, res.setting.Rek_Korpri, potKorpri, jmlDiterima ];
            
            r.eachCell({ includeEmpty: true }, (cell, colN) => { 
                cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; 
                cell.alignment = { vertical: 'middle', horizontal: colN === 2 ? 'left' : (colN === 3 || colN === 5 ? 'center' : 'right') }; 
                if(colN === 4 || colN === 6 || colN === 7) cell.numFmt = '#,##0'; 
                if (cell.value === undefined) cell.value = 0; 
            });
            currentRow++; progressCount++; updateProgress(progressCount, totalData); 
            if (progressCount % 10 === 0 || progressCount === totalData) { await new Promise(resolve => setTimeout(resolve, 0)); }
        }

        let rGrand = sheet.getRow(currentRow); rGrand.height = 30; 
        rGrand.getCell('B').value = "TOTAL KESELURUHAN"; rGrand.getCell('B').font = {bold: true};
        ['D', 'F', 'G'].forEach(col => { rGrand.getCell(col).value = { formula: `SUM(${col}7:${col}${currentRow-1})` }; });
        rGrand.eachCell({ includeEmpty: true }, (cell, colN) => { 
            cell.font = {bold: true}; 
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F3F3F3' } }; 
            cell.border = { top: {style:'medium'}, left: {style:'thin'}, bottom: {style:'medium'}, right: {style:'thin'} }; 
            if(colN === 4 || colN === 6 || colN === 7) cell.numFmt = '#,##0'; 
        });

        cetakTTDAman(sheet, currentRow + 3, res.setting, 7, res.unitCetak);
        unduhFile(wb, `Rekening_TPP_${res.bulanBesar}_${res.unitCetak}.xlsx`);
    } catch (error) { 
        Swal.close(); 
        alertError("Gagal merakit Excel: " + error.message); 
    }
}

// ==============================================================
// 4. EXCEL REKAP PAJAK TER
// ==============================================================
async function buatExcelRekapPajakJS(res) {
    try {
        const wb = new ExcelJS.Workbook();
        const sheet = wb.addWorksheet('Rekap Pajak TER', { pageSetup: { paperSize: 5, orientation: 'landscape' } });

        let cols = [ {width: 5}, {width: 22}, {width: 30}, {width: 30}, {width: 12}, {width: 10}, {width: 12}, {width: 18}, {width: 10}, {width: 15} ];
        for(let i=11; i<=15; i++) cols.push({width: 16}); cols.push({width: 12}, {width: 10}); for(let i=18; i<=20; i++) cols.push({width: 16});
        sheet.columns = cols;

        sheet.mergeCells('A1:T1'); sheet.getCell('A1').value = "REKAPITULASI PAJAK TER PENGHASILAN PEGAWAI";
        sheet.mergeCells('A2:T2'); sheet.getCell('A2').value = `PERIODE: ${res.bulanBesar}`;
        for(let i=1; i<=2; i++) { sheet.getCell(`A${i}`).font = { bold: true, size: 12 }; sheet.getCell(`A${i}`).alignment = { horizontal: 'center', vertical: 'middle' }; }

        let rowH = sheet.getRow(4); rowH.height = 55;
        rowH.values = [ "No.", "NIP", "Nama Pegawai", "Unit Kerja", "Masa\nPajak", "Tahun\nPajak", "Status\nPegawai", "NPWP", "Status\nKawin", "Jenis\nJabatan", "Gaji Bruto\n(Amprah)", "PPh 21\nGaji", "Gaji Bruto\nBersih PPh", "TPP\nBruto", "Penghasilan Kotor\n(DPP TER)", "Kategori\nTER", "Tarif\n(%)", "Total PPh\nTerutang", "PPh 21\nAmprah Gaji", "Pajak TER\n(TPP)" ];
        rowH.eachCell({ includeEmpty: true }, cell => { cell.font = { bold: true }; cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FCE5CD' } }; cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; });

        res.data.sort((a, b) => { if (a.unitKerjaLengkap < b.unitKerjaLengkap) return -1; if (a.unitKerjaLengkap > b.unitKerjaLengkap) return 1; return b.gapokRaw - a.gapokRaw; });

        let no = 1; let currentRow = 5; let progressCount = 0; let totalData = res.data.length;

        for (let c of res.data) {
            let isPPPK = String(c.golonganAsli).toUpperCase().includes("IX") || String(c.golonganAsli).toUpperCase().includes("X") || String(c.golonganAsli).toUpperCase().includes("PPPK");
            let statusPegawai = isPPPK ? "PPPK" : "PNS"; let gajiBrutoBersihPph = c.gajiKotorTER - c.pphGajiTER;
            let r = sheet.getRow(currentRow); r.height = 40; 
            r.values = [ no++, c.nip, c.nama, c.unitKerjaLengkap, res.namaBulan, res.tahun, statusPegawai, "", c.statusTER, c.jabatan, c.gajiKotorTER, c.pphGajiTER, gajiBrutoBersihPph, c.tppBruto, c.dasarPajakTER, "TER " + c.katTER, c.pctTER + "%", c.pph21TotalSebulanTER, c.pphGajiTER, c.pph21TKD ];
            r.eachCell({ includeEmpty: true }, (cell, colN) => { cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; cell.alignment = { vertical: 'middle', wrapText: true, horizontal: colN <= 4 ? 'left' : (colN <= 10 || colN === 16 || colN === 17 ? 'center' : 'right') }; if((colN >= 11 && colN <= 15) || colN >= 18) cell.numFmt = '#,##0'; if(cell.value === undefined) cell.value = ""; });
            currentRow++; progressCount++; updateProgress(progressCount, totalData); 
            if (progressCount % 10 === 0 || progressCount === totalData) { await new Promise(resolve => setTimeout(resolve, 0)); }
        }

        let rGrand = sheet.getRow(currentRow); rGrand.height = 30; sheet.mergeCells(`B${currentRow}:D${currentRow}`); rGrand.getCell('B').value = "TOTAL KESELURUHAN"; 
        ['K','L','M','N','O','R','S','T'].forEach(col => { rGrand.getCell(col).value = { formula: `SUM(${col}5:${col}${currentRow-1})` }; });
        rGrand.eachCell({ includeEmpty: true }, (cell, colN) => { cell.font = {bold: true}; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FCE5CD' } }; cell.border = { top: {style:'medium'}, left: {style:'thin'}, bottom: {style:'medium'}, right: {style:'thin'} }; cell.alignment = { vertical: 'middle', horizontal: colN <= 4 ? 'center' : 'right' }; if((colN >= 11 && colN <= 15) || colN >= 18) cell.numFmt = '#,##0'; });

        cetakTTDAman(sheet, currentRow + 3, res.setting, 20, res.unitCetak);
        unduhFile(wb, `Rekap_PajakTER_${res.bulanBesar}_${res.unitCetak}.xlsx`);
    } catch (error) { Swal.close(); alertError("Gagal merakit Excel: " + error.message); }
}

// ==============================================================
// 5. EXCEL PERHITUNGAN KOLEKTIF 
// ==============================================================
async function buatExcelPerhitunganJS(res) {
    try {
        const wb = new ExcelJS.Workbook();
        const sheet = wb.addWorksheet('Perhitungan TPP', { pageSetup: { paperSize: 5, orientation: 'landscape' } });

        let cols = [{ width: 5 }, { width: 35 }, { width: 8 }, { width: 8 }]; 
        for(let i=0; i<14; i++) cols.push({ width: 5 }); cols.push({ width: 10 }, { width: 10 }); 
        for(let i=0; i<42; i++) cols.push({ width: 12 }); cols.push({ width: 15 }); 
        sheet.columns = cols;

        sheet.mergeCells('A1:BK1'); sheet.getCell('A1').value = `PERHITUNGAN TPP ( TAMBAHAN PENGHASILAN PEGAWAI ) OPD ${res.setting.Nama_Dinas} BULAN ${res.bulanBesar}`;
        sheet.mergeCells('A2:BK2'); sheet.getCell('A2').value = `${res.jenisASN} ${res.unitCetak}`;
        for(let i=1; i<=2; i++) { sheet.getCell(`A${i}`).font = { bold: true, size: 12 }; sheet.getCell(`A${i}`).alignment = { horizontal: 'center' }; }

        sheet.mergeCells('A5:A7'); sheet.getCell('A5').value = "No."; sheet.mergeCells('B5:B7'); sheet.getCell('B5').value = "Nama / NIP / Gol. / Jabatan"; sheet.mergeCells('C5:C7'); sheet.getCell('C5').value = "Kelas\nJab."; sheet.mergeCells('D5:D7'); sheet.getCell('D5').value = "Hari\nKerja";
        sheet.mergeCells('E5:R5'); sheet.getCell('E5').value = "Keterangan Tidak Hadir (KTH)"; sheet.mergeCells('E6:E7'); sheet.getCell('E6').value = "DL"; sheet.mergeCells('F6:F7'); sheet.getCell('F6').value = "S"; sheet.mergeCells('G6:G7'); sheet.getCell('G6').value = "C"; sheet.mergeCells('H6:H7'); sheet.getCell('H6').value = "KP";
        sheet.mergeCells('I6:L6'); sheet.getCell('I6').value = "Terlambat Datang"; sheet.getCell('I7').value = "TL1"; sheet.getCell('J7').value = "TL2"; sheet.getCell('K7').value = "TL3"; sheet.getCell('L7').value = "TL4";
        sheet.mergeCells('M6:P6'); sheet.getCell('M6').value = "Cepat Pulang"; sheet.getCell('M7').value = "CP1"; sheet.getCell('N7').value = "CP2"; sheet.getCell('O7').value = "CP3"; sheet.getCell('P7').value = "CP4";
        sheet.mergeCells('Q6:Q7'); sheet.getCell('Q6').value = "TK"; sheet.mergeCells('R6:R7'); sheet.getCell('R6').value = "AS/UB"; sheet.mergeCells('S5:S7'); sheet.getCell('S5').value = "Predikat\nKerja"; sheet.mergeCells('T5:T7'); sheet.getCell('T5').value = "Tidak\nMenilai";
        
        sheet.mergeCells('U5:Z6'); sheet.getCell('U5').value = "Kriteria TPP Dasar (100%)"; ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(20 + i) + '7').value = txt);
        sheet.mergeCells('AA5:AR5'); sheet.getCell('AA5').value = "Perhitungan Produktivitas Kerja (PK 60%)"; sheet.mergeCells('AA6:AF6'); sheet.getCell('AA6').value = "Nilai PK (Setelah Bobot)"; ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(26 + i) + '7').value = txt);
        sheet.mergeCells('AG6:AL6'); sheet.getCell('AG6').value = "Potongan Absen PK"; ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(32 + i) + '7').value = txt);
        sheet.mergeCells('AM6:AR6'); sheet.getCell('AM6').value = "Akhir PK Netto"; ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(38 + i) + '7').value = txt);
        sheet.mergeCells('AS5:BJ5'); sheet.getCell('AS5').value = "Perhitungan Disiplin Kerja (DK 40%)"; sheet.mergeCells('AS6:AX6'); sheet.getCell('AS6').value = "Bobot DK Dasar (40%)"; ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(44 + i) + '7').value = txt);
        sheet.mergeCells('AY6:BD6'); sheet.getCell('AY6').value = "Potongan Absen DK"; ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(50 + i) + '7').value = txt);
        sheet.mergeCells('BE6:BJ6'); sheet.getCell('BE6').value = "Akhir DK Netto"; ['BK','PK','KK','TB','KP','Total'].forEach((txt, i) => sheet.getCell(numToLet(56 + i) + '7').value = txt);
        sheet.mergeCells('BK5:BK7'); sheet.getCell('BK5').value = "TPP Diterima\nSebelum Pajak";

        for(let i=1; i<=63; i++) { sheet.getCell(numToLet(i-1) + '8').value = i.toString(); }

        for (let i = 5; i <= 8; i++) {
            sheet.getRow(i).eachCell({ includeEmpty: true }, (cell, colN) => {
                cell.font = { bold: i!==8, italic: i===8, size: i===8 ? 8 : 9 }; cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                if(i===8) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E9ECEF' } }; else if(colN >= 5 && colN <= 18) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9EAD3' } }; else if(colN >= 27 && colN <= 44) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FCE5CD' } }; else if(colN >= 45 && colN <= 62) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E4D7F5' } }; else cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'CFE2F3' } }; 
            });
        }
        sheet.getRow(5).height = 25; sheet.getRow(6).height = 25; sheet.getRow(7).height = 30; sheet.getRow(8).height = 15;

        let grouped = {}; res.data.forEach(c => { let golDasar = String(c.golonganAsli).split("/")[0].trim().toUpperCase(); let groupName = (res.jenisASN === "PPPK") ? getGroupPPPK(golDasar) : (golDasar.includes("IX") || golDasar.includes("X") ? "PPPK" : golDasar); if(!grouped[groupName]) grouped[groupName] = []; grouped[groupName].push(c); });

        let urutanGol = res.jenisASN === "PPPK" ? ["XVII - XIII", "XII - IX", "VIII - V", "IV - I"] : ["IV", "III", "II", "I"];
        let noUrut = 1; let currentRow = 9; let subTotalRows = []; let progressCount = 0; let totalData = res.data.length; let deteksiKolomNol = Array(64).fill(0);

        for (let gol of urutanGol) {
            if(grouped[gol] && grouped[gol].length > 0) {
                let startRowGroup = currentRow;
                for (let c of grouped[gol]) {
                    let a = c.rawAbsen; let kr = [c.bk, c.pk, c.kk, c.tb, c.kp, c.tppBruto]; let nSKP = kr.map(v => v * 0.6 * c.rasioSKP); let pAbsPK = nSKP.map(v => v * c.pctAbsPK); let akhPK = nSKP.map((v, i) => v - pAbsPK[i]); let bDK = kr.map(v => v * 0.4); let pAbsDK = bDK.map(v => v * c.pctAbsDK); let akhDK = bDK.map((v, i) => v - pAbsDK[i]);
                    let rValues = [ noUrut++, `${c.nama}\nNIP. ${c.nip}\n${res.jenisASN} - Gol. ${c.golonganAsli}\n${c.jabatan}`, c.kelasJabatan || "-", c.hk, a.dl, a.s, a.c, a.kp, a.tl1, a.tl2, a.tl3, a.tl4, a.cp1, a.cp2, a.cp3, a.cp4, a.tk, a.asub, c.skp, c.menilai === "Tidak Menilai" ? "Tidak Menilai" : "", ...kr, ...nSKP, ...pAbsPK, ...akhPK, ...bDK, ...pAbsDK, ...akhDK, (akhPK[5] + akhDK[5]) ];

                    let row = sheet.getRow(currentRow); row.height = 85; row.values = rValues;
                    for(let k = 20; k < 63; k++) { deteksiKolomNol[k + 1] += parseFloat(rValues[k]) || 0; }

                    row.eachCell({ includeEmpty: true }, (cell, colN) => {
                        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; cell.alignment = { vertical: 'middle', wrapText: true, horizontal: colN === 2 ? 'left' : (colN <= 20 ? 'center' : 'right') }; if(colN >= 21) cell.numFmt = '#,##0'; if(cell.value === undefined && colN >= 21) cell.value = 0;
                    });
                    currentRow++; progressCount++; updateProgress(progressCount, totalData); 
                    if (progressCount % 5 === 0 || progressCount === totalData) { await new Promise(r => setTimeout(r, 10)); }
                }

                let rowSub = sheet.getRow(currentRow); rowSub.height = 25; rowSub.getCell('B').value = `TOTAL GOLONGAN ${gol}`; rowSub.getCell('B').font = { bold: true };
                for(let i=21; i<=63; i++) { rowSub.getCell(i).value = { formula: `SUM(${numToLet(i-1)}${startRowGroup}:${numToLet(i-1)}${currentRow-1})` }; }
                rowSub.eachCell({ includeEmpty: true }, (cell, colN) => { cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F3F3F3' } }; if (colN >= 21) { cell.numFmt = '#,##0'; cell.font = {bold: true}; } });
                subTotalRows.push(currentRow); currentRow++;
            }
        }

        let rowGrand = sheet.getRow(currentRow); rowGrand.height = 30; rowGrand.getCell('B').value = "TOTAL KESELURUHAN"; rowGrand.getCell('B').font = { bold: true };
        if(subTotalRows.length > 0) { for(let i=21; i<=63; i++) { rowGrand.getCell(i).value = { formula: subTotalRows.map(rn => `${numToLet(i-1)}${rn}`).join('+') }; } }
        rowGrand.eachCell({ includeEmpty: true }, (cell, colN) => { cell.border = { top: {style:'medium'}, left: {style:'thin'}, bottom: {style:'medium'}, right: {style:'thin'} }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E9ECEF' } }; if (colN >= 21) { cell.numFmt = '#,##0'; cell.font = {bold: true}; } });

        sheet.getColumn(1).width = 5; sheet.getColumn(2).width = 35; sheet.getColumn(3).width = 8; sheet.getColumn(4).width = 8;
        for(let i=5; i<=18; i++) sheet.getColumn(i).width = 5; sheet.getColumn(19).width = 10; sheet.getColumn(20).width = 10; 
        for(let i=21; i<=63; i++) { if (deteksiKolomNol[i] === 0) sheet.getColumn(i).width = 3.5; else sheet.getColumn(i).width = 12; }
        sheet.getColumn(63).width = 15;

        cetakTTDAman(sheet, currentRow + 3, res.setting, 63, res.unitCetak); 
        unduhFile(wb, `Perhitungan_TPP_${res.bulanBesar}_${res.unitCetak}.xlsx`);
    } catch (error) { Swal.close(); alertError("Gagal merakit Excel: " + error.message); }
}

// ==============================================================
// 6. MESIN CETAK PDF KOLEKTIF (JSPDF + AUTOTABLE + MULTI UNIT TTD)
// ==============================================================
function buatPdfNominatifJS(res) { prosesCetakPdfKolektif(res, 'Nominatif'); }
function buatPdfPerhitunganJS(res) { prosesCetakPdfKolektif(res, 'Perhitungan'); }
function buatPdfRekapGolonganJS(res) { prosesCetakPdfKolektif(res, 'RekapGolongan'); }
function buatPdfRekeningJS(res) { prosesCetakPdfKolektif(res, 'Rekening'); }
function buatPdfRekapPajakJS(res) { prosesCetakPdfKolektif(res, 'PajakTER'); }

async function prosesCetakPdfKolektif(res, jenis) {
    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({ orientation: 'landscape', format: 'legal' });

        doc.setFontSize(12); doc.setFont("helvetica", "bold");
        doc.text(`LAPORAN ${jenis.toUpperCase()} TPP - ${res.setting.Nama_Dinas}`, 14, 15);
        doc.setFontSize(10); doc.setFont("helvetica", "normal");
        doc.text(`Periode: ${res.bulanBesar} | Kategori ASN: ${res.jenisASN} ${res.unitCetak}`, 14, 21);

        let head = []; let body = []; let fRp = (num) => Math.round(num || 0).toLocaleString('id-ID');

        if (jenis === 'Rekening') {
            head = [['No', 'Nama Pegawai', 'NIP', 'No. Rekening', 'TPP Bersih', 'Pot. Korpri', 'Diterima']];
            let potKorpri = parseFloat(res.setting.Pot_Korpri) || 0;
            body = res.data.map((c, i) => [ i+1, c.nama, c.nip, c.rekening, fRp(c.tppBersih), fRp(potKorpri), fRp(c.tppBersih - potKorpri) ]);
        } else if (jenis === 'Nominatif') {
            head = [['No', 'Nama / NIP', 'Jabatan', 'Gol.', 'TPP Bruto', 'BPJS 4%', 'IWP 1%', 'PPh 21', 'Jml Potongan', 'TPP Bersih']];
            body = res.data.map((c, i) => [ i+1, `${c.nama}\n${c.nip}`, c.jabatan, c.golonganAsli, fRp(c.tppBruto), fRp(c.bpjs4), fRp(c.iwp1), fRp(c.pph21TKD), fRp(c.bpjs4 + c.iwp1 + c.pph21TKD), fRp(c.tppBersih) ]);
        } else if (jenis === 'PajakTER') {
            head = [['No', 'Nama / NIP', 'Status Kawin', 'Gaji Kotor', 'TPP Bruto', 'Dasar Pajak (TER)', 'Tarif TER', 'Pajak TPP (Potongan)']];
            body = res.data.map((c, i) => [ i+1, `${c.nama}\n${c.nip}`, c.statusTER, fRp(c.gajiKotorTER), fRp(c.tppBruto), fRp(c.dasarPajakTER), `${c.pctTER}% (TER ${c.katTER})`, fRp(c.pph21TKD) ]);
        } else {
            head = [['No', 'Nama / NIP', 'Golongan', 'Jabatan', 'Total TPP Bruto', 'Total Potongan', 'TPP Bersih Diterima']];
            body = res.data.map((c, i) => [ i+1, `${c.nama}\n${c.nip}`, c.golonganAsli, c.jabatan, fRp(c.tppBruto), fRp(c.bpjs4 + c.iwp1 + c.pph21TKD), fRp(c.tppBersih) ]);
        }

        doc.autoTable({
            startY: 25, head: head, body: body, theme: 'grid',
            headStyles: { fillColor: [41, 128, 185], textColor: 255, fontSize: 9, halign: 'center' },
            bodyStyles: { fontSize: 8 }, columnStyles: { 0: { halign: 'center', cellWidth: 10 } },
            didParseCell: function (data) { if (data.section === 'body' && data.column.index > 2) { if (/[0-9]/.test(data.cell.text[0])) { data.cell.styles.halign = 'right'; } } }
        });

        // MESIN PENCARI NAMA TTD MULTI-UNIT KHUSUS PDF
        const getVal = (key) => {
            if (res.unitCetak && res.setting[`${res.unitCetak}_${key}`]) return res.setting[`${res.unitCetak}_${key}`];
            return res.setting[key];
        };

        let finalY = doc.lastAutoTable.finalY + 15;
        if (finalY > 170) { doc.addPage(); finalY = 20; } 
        
        doc.setFontSize(9); doc.text(`Jambi, ........................ ${new Date().getFullYear()}`, 280, finalY, { align: 'center' });
        doc.text("Mengetahui,", 70, finalY + 5, { align: 'center' });
        doc.text(getVal('Kepala_Jabatan') || "KEPALA DINAS", 70, finalY + 10, { align: 'center' });
        doc.text("BENDAHARA PENGELUARAN", 280, finalY + 10, { align: 'center' });

        doc.setFont("helvetica", "bold");
        doc.text(getVal('Kepala_Nama') || "Nama Kepala", 70, finalY + 35, { align: 'center' });
        doc.text(getVal('Bendahara_Nama') || "Nama Bendahara", 280, finalY + 35, { align: 'center' });
        
        doc.setFont("helvetica", "normal");
        doc.text(`NIP. ${getVal('Kepala_NIP') || "-"}`, 70, finalY + 40, { align: 'center' });
        doc.text(`NIP. ${getVal('Bendahara_NIP') || "-"}`, 280, finalY + 40, { align: 'center' });

        doc.save(`PDF_${jenis}_${res.bulanBesar}_${res.unitCetak}.pdf`);
        Swal.close(); alertSukses(`File PDF ${jenis} Berhasil Diunduh!`);

    } catch(err) { Swal.close(); alertError("Gagal menyusun PDF: " + err.message); }
}
