// =========================================================================
// 1. KONFIGURASI SUPER MASTER (CUKUP 1 URL UNTUK SELURUH PROVINSI)
// =========================================================================
// Masukkan URL hasil Deploy Super Master Anda di sini:
const API_URL = "https://script.google.com/macros/s/AKfycbxB9FsN4pQyYenPleD-iu7Tp0G5x_vfBuip9lXJGxAXBMPcYJ4inJMypfOnkAdDB3mS/exec"; 

let listOPD = []; // Dikosongkan, karena akan ditarik otomatis dari Super Master

// =========================================================================
// 2. FUNGSI JEMBATAN API (SELALU MENGIRIMKAN KUNCI OPD)
// =========================================================================
async function fetchAPI(actionName, payloadData) {
  let token = sessionStorage.getItem('authToken') || "";
  // Ambil ID Spreadsheet OPD yang dipilih user di halaman awal
  let targetSheetId = sessionStorage.getItem('targetSheetId') || ""; 

  // Cegah request jika OPD belum dipilih (Kecuali saat ngambil daftar OPD di awal)
  if(!targetSheetId && actionName !== "getDaftarOPD") {
      return {status: "error", pesan: "Pilih OPD terlebih dahulu!"};
  }

  try {
    let response = await fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ 
          action: actionName, 
          payload: payloadData, 
          token: token,
          sheetId: targetSheetId // INI KUNCI UTAMANYA!
      })
    });
    return await response.json();
  } catch (error) { 
    return {status: "error", pesan: "Gagal terhubung ke server: " + error.message}; 
  }
}

// =========================================================================
// 3. VARIABEL GLOBAL APLIKASI
// =========================================================================
let globalBulanAktif = "", globalHariKerja = 0, globalHariKerja6 = 0, globalStatusLock = "Buka", globalPotKorpri = 0, globalRefBulanGaji = "", globalJenisASN = "", asNIPAktif = "", statusTERAktif = "", objNominatifSetahun = null, baseTPP = {bk: 0, pk: 0, kk: 0, tb: 0, kp: 0, total: 0}, arrayPeriode = [], arrayUnitKerjaValid = [], arrayJabatanValid = [], isGajiTersimpan = false, isAbsenTersimpan = false;
let currentUser = { role: "", unitkerja: "", username: "", email: "", uuid: "" };

// ==========================================
// MESIN FORMAT TITIK RIBUAN (AUTO CURRENCY)
// ==========================================
function formatRupiah(angka) {
    let number_string = String(angka).replace(/[^,\d]/g, '').toString();
    let split = number_string.split(',');
    let sisa = split[0].length % 3;
    let rupiah = split[0].substr(0, sisa);
    let ribuan = split[0].substr(sisa).match(/\d{3}/gi);
    if (ribuan) {
        let separator = sisa ? '.' : '';
        rupiah += separator + ribuan.join('.');
    }
    return split[1] != undefined ? rupiah + ',' + split[1] : rupiah;
}

function unformatRupiah(rupiah) {
    return parseInt(String(rupiah).replace(/[^0-9]/g, '')) || 0;
}

// =========================================================================
// 4. SAAT APLIKASI PERTAMA KALI DIBUKA (AUTO-FETCH DAFTAR OPD)
// =========================================================================
window.onload = async () => { 
  let savedSheetId = sessionStorage.getItem('targetSheetId');
  let isLoggedIn = sessionStorage.getItem('isLoggedIn') === 'true';

  if (savedSheetId && isLoggedIn) { 
      inisialisasiAplikasi(); 
  } else { 
      document.querySelectorAll('#viewLogin, #viewPilihBulan, #viewDaftarPegawai, #viewMasterPergub, #viewManajemenASN, #viewSetting, #viewPanduan, #viewManajemenAkun, #appContainer, #mobileHeader, #mobileBottomNav').forEach(el => el.classList.add('hidden')); 
      document.body.classList.remove('bg-gradient-login'); 
      document.getElementById('viewLanding').classList.remove('hidden'); 
      
      // TARIK DAFTAR OPD DARI SUPER MASTER SECARA DINAMIS
      startLoading("Menghubungkan ke Server Provinsi...");
      let res = await fetchAPI("getDaftarOPD", {});
      stopLoading();

      let opdSelect = document.getElementById('selectTenantOPD');
      opdSelect.innerHTML = '<option value="">-- PILIH OPD --</option>';

      if (res && res.status === "sukses") {
          listOPD = res.data;
          listOPD.forEach(opd => { 
              // Simpan ID Spreadsheet ke dalam value dropdown
              opdSelect.innerHTML += `<option value="${opd.sheetId}">${opd.nama}</option>`; 
          });
      } else {
          opdSelect.innerHTML = '<option value="">Gagal memuat OPD</option>';
          alertError("Gagal mengambil daftar OPD: " + (res.pesan || "Cek koneksi internet"));
      }
  }
};

// =========================================================================
// 5. NAVIGASI LANDING & LOGIN
// =========================================================================
function ubahNamaOPD(element) { 
  if(element.value !== "") {
      sessionStorage.setItem('targetSheetId', element.value); 
  } else {
      sessionStorage.removeItem('targetSheetId');
  }
}

async function masukKeLogin() {
  let target = sessionStorage.getItem('targetSheetId');
  if(!target) return alertPeringatan("Silakan pilih Instansi / OPD Anda terlebih dahulu!");
  
  startLoading("Menghubungkan ke Database OPD..."); 
  await inisialisasiAplikasi(); 
  stopLoading();
  
  document.getElementById('viewLanding').classList.add('hidden'); 
  document.getElementById('viewLogin').classList.remove('hidden'); 
  document.body.classList.add('bg-gradient-login');
}

function kembaliKeLanding() {
  sessionStorage.clear(); 
  document.getElementById('selectTenantOPD').value = "";
  
  let appCont = document.getElementById('appContainer');
  if(appCont) appCont.classList.add('hidden');
  
  let mHeader = document.getElementById('mobileHeader');
  if(mHeader) mHeader.classList.add('hidden');
  
  let mNav = document.getElementById('mobileBottomNav');
  if(mNav) mNav.classList.add('hidden');
  
  document.getElementById('viewLogin').classList.add('hidden'); 
  document.getElementById('viewLanding').classList.remove('hidden'); 
  document.body.classList.remove('bg-gradient-login');
}

function togglePassword(inputId, btn) { 
  let inp = document.getElementById(inputId); let icon = btn.querySelector('i'); 
  if(inp.type === "password") { inp.type = "text"; icon.classList.replace('bi-eye', 'bi-eye-slash'); } 
  else { inp.type = "password"; icon.classList.replace('bi-eye-slash', 'bi-eye'); } 
}

async function doLogin(e) {
  e.preventDefault(); startLoading("Memeriksa Akses..."); 
  let u = document.getElementById('logUser').value; 
  let p = document.getElementById('logPass').value;
  
  let res = await fetchAPI("prosesLogin", {username: u, password: p});
  stopLoading(); 
  
  if(res.status === "sukses") {
    sessionStorage.setItem('authToken', res.token); currentUser = res.user; 
    document.getElementById('navInfoRole').innerText = currentUser.unitkerja; 
    document.getElementById('mHeaderUnit').innerText = currentUser.unitkerja; 
    
    let r = currentUser.role;
      let isSuper = (r === "Super Admin");
      let isOpd = (r === "Admin OPD");

      // 1. MASTER TPP: Hanya Super Admin
      if(isSuper) { 
          document.getElementById('btnMasterAdmin').classList.remove('hidden'); 
      } else { 
          document.getElementById('btnMasterAdmin').classList.add('hidden'); 
      }
      
      // 2. KELOLA AKUN: Super Admin DAN Admin OPD bisa lihat!
      if(isSuper || isOpd) { 
          document.getElementById('btnAkunAdmin').classList.remove('hidden'); 
      } else { 
          document.getElementById('btnAkunAdmin').classList.add('hidden'); 
      }
      
      // 3. SETTING: Hanya Super Admin & Admin OPD (Operator disembunyikan)
      if(isSuper || isOpd) { 
          document.getElementById('btnSettingAdmin').classList.remove('hidden'); 
          if(document.getElementById('mHeaderAdminIcons')) document.getElementById('mHeaderAdminIcons').classList.remove('hidden');
      } else { 
          document.getElementById('btnSettingAdmin').classList.add('hidden'); 
          if(document.getElementById('mHeaderAdminIcons')) document.getElementById('mHeaderAdminIcons').classList.add('hidden');
      }
      
      // 4. IKON DI HP MOBILE
      let iconMasterHP = document.querySelector('i[onclick="switchView(\'viewMasterPergub\')"]');
      let iconAkunHP = document.querySelector('i[onclick="muatDaftarAkun()"]');

      if(isSuper || isOpd) { 
          document.getElementById('btnTambahBulan').classList.remove('hidden'); 
          document.getElementById('btnKelolaBulan').classList.remove('hidden'); 
          
          // Ikon Master TPP HP: Hanya Super Admin
          if(iconMasterHP && isSuper) iconMasterHP.classList.remove('hidden'); 
          else if(iconMasterHP) iconMasterHP.classList.add('hidden');
          
          // Ikon Kelola Akun HP: Super Admin & Admin OPD
          if(iconAkunHP && (isSuper || isOpd)) iconAkunHP.classList.remove('hidden'); 
          else if(iconAkunHP) iconAkunHP.classList.add('hidden');
      } else { 
          document.getElementById('btnTambahBulan').classList.add('hidden'); 
          document.getElementById('btnKelolaBulan').classList.add('hidden'); 
          if(iconMasterHP) iconMasterHP.classList.add('hidden');
          if(iconAkunHP) iconAkunHP.classList.add('hidden');
      }

    document.getElementById('viewLogin').classList.add('hidden'); 
    document.getElementById('mainNav').classList.remove('hidden');
    sessionStorage.setItem('isLoggedIn', 'true'); 
    sessionStorage.setItem('currentUser', JSON.stringify(currentUser));
    switchView('viewPilihBulan'); 
    inisialisasiAplikasi();
  } else { 
    alertError(res.pesan); 
  }
}

  function logoutApp() { 
    Swal.fire({ 
        title: 'Yakin ingin keluar?', 
        text: "Anda akan dikembalikan ke halaman awal.", 
        icon: 'question', 
        showCancelButton: true, 
        confirmButtonColor: '#dc3545', 
        cancelButtonColor: '#6c757d', 
        confirmButtonText: '<i class="bi bi-box-arrow-right"></i> Ya, Keluar!', 
        cancelButtonText: 'Batal' 
    }).then((result) => {
      if (result.isConfirmed) {
        kembaliKeLanding(); 
        currentUser = { role: "", unitkerja: "", username: "", email: "", uuid: "" }; 
        globalBulanAktif = ""; globalHariKerja = 0; globalHariKerja6 = 0; globalStatusLock = "Buka";
        document.getElementById('appContainer').classList.add('hidden'); 
        document.getElementById('mobileHeader').classList.add('hidden'); 
        document.getElementById('mobileBottomNav').classList.add('hidden'); 
        document.getElementById('logUser').value = ""; 
        document.getElementById('logPass').value = ""; 
        toggleTombolPegawaiHP(false); 
        alertSukses("Anda berhasil logout dari sistem.");
      }
    });
  }

  function switchView(viewId) { 
    // PERBAIKAN: #viewManajemenASN dihapus agar tidak bentrok dengan Popup
    document.querySelectorAll('#viewLogin, #viewLanding, #viewPilihBulan, #viewDaftarPegawai, #viewMasterPergub, #viewSetting, #viewPanduan, #viewManajemenAkun').forEach(el => el.classList.add('hidden')); 
    
    let appContainer = document.getElementById('appContainer'); 
    let mHeader = document.getElementById('mobileHeader'); 
    let mBottomNav = document.getElementById('mobileBottomNav');

    if(viewId === 'viewLogin') { 
        document.body.classList.add('bg-gradient-login'); 
        if(appContainer) appContainer.classList.add('hidden'); 
        if(mHeader) mHeader.classList.add('hidden'); 
        if(mBottomNav) mBottomNav.classList.add('hidden'); 
        document.getElementById('viewLogin').classList.remove('hidden'); 
    } 
    else if (viewId === 'viewLanding') { 
        document.body.classList.remove('bg-gradient-login'); 
        if(appContainer) appContainer.classList.add('hidden'); 
        if(mHeader) mHeader.classList.add('hidden'); 
        if(mBottomNav) mBottomNav.classList.add('hidden'); 
        document.getElementById('viewLanding').classList.remove('hidden'); 
    }
    else { 
        document.body.classList.remove('bg-gradient-login'); 
        document.body.style.backgroundColor = "#f4f6f9"; 
        if(appContainer) appContainer.classList.remove('hidden'); 
        if(mHeader) mHeader.classList.remove('hidden'); 
        if(mBottomNav) mBottomNav.classList.remove('hidden'); 
        document.getElementById(viewId).classList.remove('hidden'); 
    }

    document.querySelectorAll('.m-nav-item').forEach(el => el.classList.remove('active'));
    if(viewId === 'viewPilihBulan') document.getElementById('mNavHome').classList.add('active'); 
    if(viewId === 'viewPanduan') document.getElementById('mNavPanduan').classList.add('active'); 
    if(viewId === 'viewDaftarPegawai' || viewId === 'viewManajemenASN') document.getElementById('mNavPegawai').classList.add('active');
    
    if(viewId !== 'viewLogin' && viewId !== 'viewLanding') { 
        sessionStorage.setItem('currentView', viewId); 
    }

    if (viewId !== 'viewDaftarPegawai' && viewId !== 'viewManajemenASN' && viewId !== 'viewLogin' && viewId !== 'viewLanding') {
        let btnMenuPegawai = document.getElementById('btnMenuPegawai');
        if (btnMenuPegawai) btnMenuPegawai.disabled = true;
        toggleTombolPegawaiHP(false);
    }
  }

  function masukAplikasi() {
    let b = document.getElementById('pilihPeriodeUtama').value; 
    if(!b) { alertPeringatan("Pilih atau Tambah periode terlebih dahulu!"); return; }
    
    globalJenisASN = document.getElementById('pilihJenisASN').value; 
    globalBulanAktif = b;
    
    let obj = arrayPeriode.find(x => x.namaPeriode === b); 
    globalHariKerja = obj ? obj.hariKerja : 22; 
    globalHariKerja6 = obj ? obj.hariKerja6 : 26; 
    globalStatusLock = obj ? obj.statusLock : "Buka"; 
    globalRefBulanGaji = obj ? (obj.refBulanGaji || b) : b; 
    globalJenisPeriode = obj ? (obj.jenisPeriode || "Reguler") : "Reguler";
    
    if(document.getElementById('labelBulanPencairan')) document.getElementById('labelBulanPencairan').innerText = globalRefBulanGaji; 
    document.getElementById('labelBulanAktif').innerText = b + " (" + globalJenisASN + ")"; 
    document.getElementById('labelHKAktif').innerText = `${globalHariKerja} / ${globalHariKerja6}`; 
    
    document.getElementById('btnMenuPegawai').disabled = false; 
    toggleTombolPegawaiHP(true);
    
    if (globalStatusLock === "Kunci" || currentUser.role === "Operator" || currentUser.role === "Admin Sub Unit/UPTD") { 
      document.getElementById('warningLock').classList.remove('hidden'); 
      if(currentUser.role === "Operator" || currentUser.role === "Admin Sub Unit/UPTD") {
         document.getElementById('warningLock').innerHTML = "<i class='bi bi-info-circle'></i> Hanya Admin OPD yang dapat menambah/import Pegawai baru.";
      } else {
         document.getElementById('warningLock').innerHTML = "<i class='bi bi-lock-fill'></i> PERIODE INI SUDAH DIKUNCI <br> <span class='small text-dark'>Data hanya dapat dilihat dan dicetak.</span>";
      }
      document.getElementById('btnTambahPegawai').classList.add('hidden'); 
      if(document.getElementById('boxImportExcel')) document.getElementById('boxImportExcel').classList.add('hidden'); 
    } else { 
      document.getElementById('warningLock').classList.add('hidden'); 
      document.getElementById('btnTambahPegawai').classList.remove('hidden'); 
    }
    
    // Simpan ke Sesi Browser
    sessionStorage.setItem('globalBulanAktif', globalBulanAktif); 
    sessionStorage.setItem('globalJenisASN', globalJenisASN); 
    updateFilterGolDropdown(); 
    
    // Pindah Tampilan
    switchView('viewDaftarPegawai'); 

    // 🧠 LOGIKA SMART CACHE: Cek apakah user pindah ke bulan yang berbeda
    if (window.cacheDataPegawaiBulan !== globalBulanAktif) {
        // Jika beda bulan: Hapus cache lama, wajib loading data baru
        window.cacheDataPegawaiAll = null;
        muatDataPegawai(true); 
    } else {
        // Jika bulan sama (hanya ganti PNS/PPPK): Jalur kilat tanpa loading!
        muatDataPegawai(false); 
    }
  }

  const escapeStr = (str) => String(str).replace(/'/g, "\\'").replace(/"/g, "&quot;");

  function updateFilterGolDropdown() {
    let sel = document.getElementById('filterGol'); let html = '<option value="">Semua Golongan</option>';
    if (globalJenisASN === "PNS") { html += '<option value="IV">Golongan IV</option><option value="III">Golongan III</option><option value="II">Golongan II</option><option value="I">Golongan I</option>'; } 
    else if (globalJenisASN === "PPPK") { html += '<option value="XVII - XIII">Golongan XVII - XIII</option><option value="XII - IX">Golongan XII - IX</option><option value="VIII - V">Golongan VIII - V</option><option value="IV - I">Golongan IV - I</option>'; }
    sel.innerHTML = html;
  }

  function bukaModalTambahPeriode() {
      let nextBulanAngka = getSmartBulanAngka();
      let nextTahun = new Date().getFullYear(); 
      
      if(arrayPeriode && arrayPeriode.length > 0) { 
          let lastReguler = [...arrayPeriode].reverse().find(p => p.jenisPeriode === "Reguler" || !p.jenisPeriode); 
          if(lastReguler && !isNaN(lastReguler.tahun)) { 
              nextTahun = parseInt(lastReguler.tahun);
              let isDesemberAda = arrayPeriode.some(p => (p.jenisPeriode === "Reguler" || !p.jenisPeriode) && parseInt(p.bulanAngka) === 12);
              if (nextBulanAngka === 1 && isDesemberAda) { 
                  nextTahun += 1; 
              } 
          } 
      } 

      const namaBulan = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
      
      document.getElementById('pTahun').value = nextTahun; 
      document.getElementById('pBulanAngka').value = nextBulanAngka; 
      document.getElementById('pBulanNama').value = namaBulan[nextBulanAngka] + " " + nextTahun; 
      document.getElementById('pHariKerja').value = ""; 
      document.getElementById('pHariKerja6').value = ""; 
      document.getElementById('pJenisPeriode').value = "Reguler"; 
      
      toggleIndukPajak();
      
      let selInduk = document.getElementById('pIndukPajak'); 

      selInduk.innerHTML = '<option value="">-- Pilih Bulan Gaji --</option>'; 
      // Looping otomatis 12 bulan untuk tahun yang sedang dipilih
      for(let i = 1; i <= 12; i++) {
    let namaPer = namaBulan[i] + " " + nextTahun;
    selInduk.innerHTML += `<option value="${namaPer}">${namaPer}</option>`;
      }

      document.getElementById('warningIndukPajak').classList.add('hidden');

      let modalObj = bootstrap.Modal.getInstance(document.getElementById('modalTambahPeriode')) || new bootstrap.Modal(document.getElementById('modalTambahPeriode'));
      modalObj.show();
  }

  function toggleIndukPajak() {
    let jenis = document.getElementById('pJenisPeriode').value;
    let inputNama = document.getElementById('pBulanNama');
    let inputAngka = document.getElementById('pBulanAngka');
    let tahun = document.getElementById('pTahun').value || new Date().getFullYear();
    let nextBulanAngka = getSmartBulanAngka(); 
    
    // --- MODIFIKASI DI SINI ---
    // Kita hapus logika yang menyembunyikan boxIndukPajak
    let boxInduk = document.getElementById('boxIndukPajak');
    boxInduk.classList.remove("hidden"); // Paksa tampil terus
    
    const namaBulan = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    
    if(jenis === "Reguler") { 
        inputNama.value = namaBulan[nextBulanAngka] + " " + tahun; 
        inputAngka.value = nextBulanAngka;
        // Biarkan user memilih gaji bulan mana yang jadi acuan
    } else { 
        inputNama.value = jenis + " " + tahun; 
        inputAngka.value = (jenis === "THR") ? 14 : 13; 
        document.getElementById('pHariKerja').value = 22; 
        document.getElementById('pHariKerja6').value = 26; 
    }
}

  function sinkronHariKerja() { let namaInduk = document.getElementById('pIndukPajak').value; if(!namaInduk) return; let dataInduk = arrayPeriode.find(p => p.namaPeriode === namaInduk); if(dataInduk) { document.getElementById('pHariKerja').value = dataInduk.hariKerja; document.getElementById('pHariKerja6').value = dataInduk.hariKerja6; } }

  // =========================================================
  // FUNGSI: TAMBAH PERIODE (SALIN/KOSONG/IMPORT) - AUTO REFRESH
  // =========================================================
  async function simpanPeriodeLanjut(e) {
      e.preventDefault();
      let btnSubmit = e.target.querySelector('button[type="submit"]');
      let sumber = document.getElementById('pSumberData').value;
      let fileInput = document.getElementById('pFileImportPegawai');

      if (sumber === 'import' && !fileInput.files[0]) {
          return alertPeringatan("Pilih file Excel BKD terlebih dahulu untuk diimpor!");
      }

      btnSubmit.disabled = true;
      let jenis = document.getElementById('pJenisPeriode').value;
      let tahunInput = document.getElementById('pTahun').value;
      let namaPeriode = document.getElementById('pBulanNama').value;
      let indukPajakVal = (jenis === "Reguler") ? "" : document.getElementById('pIndukPajak').value;

      if (jenis === "THR" || jenis === "Gaji 13") {
          let sudahAda = arrayPeriode.find(p => p.jenisPeriode === jenis && p.tahun == tahunInput);
          if (sudahAda) {
              btnSubmit.disabled = false;
              return alertPeringatan(`Gagal! Periode ${jenis} untuk tahun ${tahunInput} sudah pernah ditambahkan.`);
          }
      }

      startLoading("Membuka Periode Baru & Menyinkronkan Data...");

      const d = {
          tahun: tahunInput, bulanAngka: document.getElementById('pBulanAngka').value, namaPeriode: namaPeriode, hariKerja: document.getElementById('pHariKerja').value, hariKerja6: document.getElementById('pHariKerja6').value,
          salinData: (sumber === 'salin'), 
          jenisPeriode: jenis, refBulanGaji: namaPeriode, indukPajak: indukPajakVal
      };

      let res = await fetchAPI("simpanPeriode", d);

      if (res.status !== "sukses") {
          stopLoading(); btnSubmit.disabled = false;
          return alertError(res.pesan);
      }

      if (sumber === 'import') {
          // Jika sumbernya Import, kita panggil fungsi import yang sudah diperbaiki
          prosesImportExcel(fileInput, namaPeriode, res.dataTerbaru, btnSubmit);
      } 
      else {
          // Jika Salin / Kosong
          stopLoading(); btnSubmit.disabled = false;
          alertSukses(res.pesan);
          
          bootstrap.Modal.getInstance(document.getElementById('modalTambahPeriode')).hide();
          
          // 1. Render ulang dropdown dengan data baru
          renderDropdownPeriode(res.dataTerbaru);
          
          // 2. Set dropdown ke bulan yang baru saja dibuat
          document.getElementById('pilihPeriodeUtama').value = namaPeriode;
          
          // 3. Hancurkan memori cache lama agar aplikasi sadar ada data baru
          window.cacheDataPegawaiAll = null; 
          globalBulanAktif = namaPeriode;
          sessionStorage.setItem('globalBulanAktif', globalBulanAktif);
          
          // 4. Paksa masuk dan tarik data terbaru dari server
          masukAplikasi(); 
      }
  }

  async function muatDaftarPeriode() { let data = await fetchAPI("getDaftarPeriode", {}); if (data && !data.error) { renderDropdownPeriode(data); } }

  function renderDropdownPeriode(data) {
    if (data && data.length > 0) {
        const getSortValue = (p) => {
            let baseBulan = parseInt(p.bulanAngka) || 0;
            let year = parseInt(p.tahun) || 0;
            if (p.jenisPeriode !== "Reguler" && p.indukPajak) {
                let induk = data.find(x => x.namaPeriode === p.indukPajak);
                if (induk) { baseBulan = parseInt(induk.bulanAngka) || 0; year = parseInt(induk.tahun) || 0; }
                if (p.jenisPeriode === "THR") baseBulan += 0.1;
                else if (p.jenisPeriode === "Gaji 13") baseBulan += 0.2;
                else baseBulan += 0.3;
            }
            return (year * 100) + baseBulan;
        };
        data.sort((a, b) => getSortValue(a) - getSortValue(b));
    }

    arrayPeriode = data; 
    let container = document.getElementById('containerTombolPeriode'); 
    let hiddenInput = document.getElementById('pilihPeriodeUtama');
    let displayTeks = document.getElementById('displayBulanAktif');
    
    container.innerHTML = "";
    
    if(!data || data.length === 0) { 
        container.innerHTML = `<div class="col-12"><div class="alert alert-warning w-100 mb-0 fw-bold"><i class="bi bi-info-circle"></i> Belum ada data bulan. Silakan klik "Tambah Periode Baru" di bawah!</div></div>`; 
        displayTeks.innerText = "KOSONG";
        hiddenInput.value = "";
    } else { 
        data.forEach(p => { 
            // PERUBAHAN 2: Deteksi gembok & berikan warna tombol yang berbeda
            let isLocked = p.statusLock === "Kunci";
            let lockIcon = isLocked ? '<i class="bi bi-lock-fill me-1"></i>' : '';
            // Jika dikunci, dasarnya abu-abu. Jika terbuka, dasarnya biru.
            let btnClass = isLocked ? 'btn-outline-secondary text-secondary' : 'btn-outline-primary';
            
            container.innerHTML += `
            <div class="col-6 col-md-3">
                <button type="button" class="btn ${btnClass} w-100 fw-bold py-2 btn-periode-select shadow-sm text-truncate" onclick="klikBulan('${p.namaPeriode}', '${p.statusLock}', this)" title="${p.namaPeriode}">
                    ${lockIcon}${p.namaPeriode}
                </button>
            </div>`; 
        }); 
        
        let selectedBulan = "";
        if (globalBulanAktif && data.some(p => p.namaPeriode === globalBulanAktif)) {
            selectedBulan = globalBulanAktif;
        } else {
            selectedBulan = data[data.length - 1].namaPeriode; 
        }
        
        let selectedBulanObj = data.find(p => p.namaPeriode === selectedBulan) || data[data.length - 1];
        
        setTimeout(() => {
            let btns = container.querySelectorAll('.btn-periode-select');
            btns.forEach(btn => {
                if (btn.innerText.replace(/[\n\r]+|[\s]{2,}/g, ' ').trim().includes(selectedBulanObj.namaPeriode)) {
                    klikBulan(selectedBulanObj.namaPeriode, selectedBulanObj.statusLock, btn);
                }
            });
        }, 50);
    }
  }

  function klikBulan(namaBulan, statusLock, btnElement) {
      document.getElementById('pilihPeriodeUtama').value = namaBulan;
      document.getElementById('displayBulanAktif').innerText = namaBulan.toUpperCase();
      
      let btns = document.querySelectorAll('.btn-periode-select');
      btns.forEach(b => {
          b.classList.remove('btn-primary', 'btn-secondary', 'text-white');
          if(b.innerHTML.includes('bi-lock-fill')) {
              b.classList.add('btn-outline-secondary', 'text-secondary');
          } else {
              b.classList.add('btn-outline-primary');
          }
      });
      
      btnElement.classList.remove('btn-outline-primary', 'btn-outline-secondary', 'text-secondary');
      if (statusLock === "Kunci") {
          btnElement.classList.add('btn-secondary', 'text-white');
      } else {
          btnElement.classList.add('btn-primary', 'text-white');
      }
  }

  function ubahJenisASN(jenis) {
      document.getElementById('pilihJenisASN').value = jenis;
      let btnPNS = document.getElementById('btnPilihPNS');
      let btnPPPK = document.getElementById('btnPilihPPPK');
      
      if (jenis === 'PNS') {
          btnPNS.classList.replace('btn-outline-success', 'btn-success');
          btnPNS.classList.add('text-white');
          btnPPPK.classList.replace('btn-warning', 'btn-outline-warning');
          btnPPPK.classList.remove('text-white');
      } else {
          btnPPPK.classList.replace('btn-outline-warning', 'btn-warning');
          btnPPPK.classList.add('text-dark');
          btnPNS.classList.replace('btn-success', 'btn-outline-success');
          btnPNS.classList.remove('text-white');
      }
  }

  function bukaModalEditPeriode() {
    let b = document.getElementById('pilihPeriodeUtama').value; if(!b) return alertPeringatan("Pilih periode terlebih dahulu!");
    let obj = arrayPeriode.find(x => x.namaPeriode === b); document.getElementById('eBulanNama').value = obj.namaPeriode; document.getElementById('eHariKerja').value = obj.hariKerja; document.getElementById('eHariKerja6').value = obj.hariKerja6; document.getElementById('eHk5Lama').value = obj.hariKerja; document.getElementById('eHk6Lama').value = obj.hariKerja6; document.getElementById('eStatusLock').value = obj.statusLock || "Buka";

    // CARI FUNGSI bukaModalEditPeriode() dan ubah bagian "let selRef" ke bawah menjadi seperti ini:
let selRef = document.getElementById('editRefBulanGaji'); 
if(selRef) { 
    selRef.innerHTML = ""; 
    let tahun = obj.tahun || b.split(" ")[1] || new Date().getFullYear();
    const namaBulan = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    for(let i=1; i<=12; i++) {
        let namaPer = namaBulan[i] + " " + tahun;
        selRef.innerHTML += `<option value="${namaPer}">${namaPer}</option>`;
    }
    selRef.value = obj.refBulanGaji || obj.namaPeriode; 
    cekBulanReferensi(selRef.value, 'warningEditRef'); // Panggil pengecekan saat modal dibuka
}
    
    let modalObj = bootstrap.Modal.getInstance(document.getElementById('modalEditPeriode')) || new bootstrap.Modal(document.getElementById('modalEditPeriode'));
    modalObj.show();
  }

  async function submitEditPeriode(e) {
    e.preventDefault(); let b = document.getElementById('eBulanNama').value; let d = { bulan: b, hk5: document.getElementById('eHariKerja').value, hk6: document.getElementById('eHariKerja6').value, hk5Lama: document.getElementById('eHk5Lama').value, hk6Lama: document.getElementById('eHk6Lama').value, statusLock: document.getElementById('eStatusLock').value, refBulan: document.getElementById('editRefBulanGaji').value };
    startLoading("Menyimpan Pengaturan Periode..."); let res = await fetchAPI("updateHariKerja", d); stopLoading(); if(res.status === "sukses") { alertSukses(res.pesan); bootstrap.Modal.getInstance(document.getElementById('modalEditPeriode')).hide(); renderDropdownPeriode(res.dataTerbaru); document.getElementById('pilihPeriodeUtama').value = b; } else { alertError(res.pesan); }
  }

  function konfirmasiHapusPeriode() {
    let b = document.getElementById('eBulanNama').value;
    Swal.fire({
      title: 'Hapus Periode Permanen?',
      html: `AWAS! Anda akan menghapus periode <b>"${b}"</b>.<br>Ini akan menghapus <b>SELURUH DATA PEGAWAI, GAJI, DAN TPP</b> pada bulan tersebut. Yakin ingin melanjutkan?`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#dc3545',
      cancelButtonColor: '#6c757d',
      confirmButtonText: '<i class="bi bi-trash"></i> Ya, Hapus Semua!',
      cancelButtonText: 'Batal'
    }).then(async (result) => {
      if (result.isConfirmed) {
        startLoading("Menghapus Periode..."); 
        let res = await fetchAPI("hapusPeriodeUtama", b); 
        stopLoading(); 
        if(res.status === "sukses") { 
          alertSukses(res.pesan); 
          bootstrap.Modal.getInstance(document.getElementById('modalEditPeriode')).hide(); 
          renderDropdownPeriode(res.dataTerbaru); 
        } else { alertError(res.pesan); }
      }
    });
  }

  async function muatDataUnitKerja() { 
    let data = await fetchAPI("getDaftarUnitKerja", {}); 
    if (data && !data.error) { arrayUnitKerjaValid = data;  

    setupAutocomplete(document.getElementById('mUnitKerja'), arrayUnitKerjaValid);
    setupAutocomplete(document.getElementById('uUnitKerjaOp'), arrayUnitKerjaValid);

    let list = document.getElementById('listUnitKerjaUtama'); 
    if(list) { list.innerHTML = ""; 
    data.forEach(unit => { list.innerHTML += `<option value="${unit}">`; }); } } }

  let globalDataPergub = []; let currentPergubPage = 1; let pergubRowsPerPage = 15;
  async function muatDataPergub() { let data = await fetchAPI("getDaftarPergub", {}); if (data && !data.error) { globalDataPergub = data; terapkanFilterPergub(); } }

  function terapkanFilterPergub() {
    let cari = document.getElementById('filterCariPergub').value.toLowerCase(); let filtered = globalDataPergub.filter(row => { return String(row[1]).toLowerCase().includes(cari); });
    filtered.sort((a, b) => { let statusA = String(a[9] || "PNS").toUpperCase(); let statusB = String(b[9] || "PNS").toUpperCase(); if (statusA === "PNS" && statusB !== "PNS") return -1; if (statusA !== "PNS" && statusB === "PNS") return 1;  return 0; });
    renderTabelPergub(filtered);
  }
  
  function renderTabelPergub(data) {
    let tbody = document.getElementById('tabelBodyPergub');
    if(!data.length) { document.getElementById('infoPaginationPergub').innerText = "Menampilkan 0 data"; document.getElementById('btnPaginationPergub').innerHTML = ""; return tbody.innerHTML = '<tr><td colspan="4" class="text-center text-danger fw-bold">Data tidak ditemukan.</td></tr>'; }
    let totalPages = Math.ceil(data.length / pergubRowsPerPage); if(currentPergubPage > totalPages) currentPergubPage = totalPages; let start = (currentPergubPage - 1) * pergubRowsPerPage; let end = start + pergubRowsPerPage; let paginatedData = data.slice(start, end); let html = "";
    paginatedData.forEach(row => { let total = parseFloat(row[8]) || 0; let jab = escapeStr(row[1]); let kls = escapeStr(row[2]); let bk=row[3], pk=row[4], kk=row[5], tb=row[6], kp=row[7]; let statusPeg = escapeStr(row[9] || "PNS"); let args = `'lihat', '${jab}', '${kls}', ${bk}, ${pk}, ${kk}, ${tb}, ${kp}, '${statusPeg}'`; let argsE = `'edit', '${jab}', '${kls}', ${bk}, ${pk}, ${kk}, ${tb}, ${kp}, '${statusPeg}'`; let badgeStatus = statusPeg === "PPPK" ? "bg-warning text-dark" : "bg-primary"; html += `<tr><td><b>${jab}</b> <span class="badge ${badgeStatus} ms-1">${statusPeg}</span></td><td>${kls}</td><td>Rp ${Math.round(total).toLocaleString('id-ID')}</td><td class="text-center"><button class="btn btn-sm btn-info text-white me-1" onclick="bukaModalPergub(${args})">Lihat</button><button class="btn btn-sm btn-primary me-1" onclick="bukaModalPergub(${argsE})">Edit</button><button class="btn btn-sm btn-danger" onclick="hapusPergub('${jab}')">Del</button></td></tr>`; });
    tbody.innerHTML = html; document.getElementById('infoPaginationPergub').innerText = `Tampil ${start + 1} - ${Math.min(end, data.length)} dari ${data.length}`; let btnHtml = `<button class="btn btn-sm btn-outline-primary" onclick="ubahHalamanPergub(${currentPergubPage - 1})" ${currentPergubPage === 1 ? 'disabled' : ''}>Mundur</button><span class="btn btn-sm btn-primary disabled text-white fw-bold">Hal ${currentPergubPage}/${totalPages}</span><button class="btn btn-sm btn-outline-primary" onclick="ubahHalamanPergub(${currentPergubPage + 1})" ${currentPergubPage === totalPages ? 'disabled' : ''}>Maju</button>`; document.getElementById('btnPaginationPergub').innerHTML = btnHtml;
  }
  function ubahHalamanPergub(page) { currentPergubPage = page; terapkanFilterPergub(); }

  function muatDropdownJabatan(selectedValue = "") {
    if (globalDataPergub && globalDataPergub.length > 0) { 
        let dataTerfilter = globalDataPergub.filter(row => String(row[9] || "PNS") === globalJenisASN); 
        arrayJabatanValid = dataTerfilter.map(r => String(r[1]).trim()); 

        let arrJabatanObj = dataTerfilter.map(row => {
            let totalTPP = parseFloat(row[8]) || 0;
            return {
                display: `${row[1]} <br><small class="text-success">Kls: ${row[2]} - Rp ${totalTPP.toLocaleString('id-ID')}</small>`,
                value: String(row[1]).trim()
            };
        });
        setupAutocomplete(document.getElementById('mJabatan'), arrJabatanObj);

        if(selectedValue) document.getElementById('mJabatan').value = selectedValue; 
    }
  }

  function bukaModalPergub(aksi, jabatan="", kelas="", bk=0, pk=0, kk=0, tb=0, kp=0, statusPegawai="PNS") {
    document.getElementById('pAksi').value = aksi; document.getElementById('pJabatanLama').value = jabatan; document.getElementById('pJabatan').value = jabatan; document.getElementById('pKelas').value = kelas; document.getElementById('pBK').value = bk; document.getElementById('pPK').value = pk; document.getElementById('pKK').value = kk; document.getElementById('pTB').value = tb; document.getElementById('pKP').value = kp; document.getElementById('pStatusPegawai').value = statusPegawai; 
    hitungTotalPergub(); let els = document.getElementById('formPergub').elements; let isRO = (aksi === 'lihat'); for (let i = 0; i < els.length; i++) { if(els[i].id !== 'btnSimpanPergub' && els[i].id !== 'pTotal') { els[i].readOnly = isRO; if(isRO) els[i].classList.add('readonly-field'); else els[i].classList.remove('readonly-field'); } } document.getElementById('pergubModalTitle').innerText = aksi === 'tambah' ? 'Tambah Master TPP' : (aksi === 'edit' ? 'Edit Master TPP' : 'Detail Master TPP'); document.getElementById('btnSimpanPergub').style.display = isRO ? 'none' : 'block'; 
    let modalObj = bootstrap.Modal.getInstance(document.getElementById('modalKelolaPergub')) || new bootstrap.Modal(document.getElementById('modalKelolaPergub'));
    modalObj.show();
  }

  async function submitPergub(e) { 
    e.preventDefault(); let aksi = document.getElementById('pAksi').value; startLoading("Menyimpan Master...");
    const d = { jabatanLama: document.getElementById('pJabatanLama').value, namaJabatan: document.getElementById('pJabatan').value, kelasJabatan: document.getElementById('pKelas').value, bk: document.getElementById('pBK').value, pk: document.getElementById('pPK').value, kk: document.getElementById('pKK').value, tb: document.getElementById('pTB').value, kp: document.getElementById('pKP').value, totalTpp: document.getElementById('pTotal').value, statusPegawai: document.getElementById('pStatusPegawai').value }; 

    let res = await fetchAPI(aksi === 'edit' ? "updatePergub" : "simpanPergub", d); 
  stopLoading(); 
  
  // PERBAIKAN DI SINI
  if(res && res.status === "sukses") { 
      alertSukses(res.pesan); 
      bootstrap.Modal.getInstance(document.getElementById('modalKelolaPergub')).hide(); 
      muatDataPergub(); 
  } else { 
      alertError(res.pesan || "Terjadi kesalahan!"); 
  }
  }

  function hitungTotalPergub() { let inputs = ['pBK', 'pPK', 'pKK', 'pTB', 'pKP'].map(id => parseFloat(document.getElementById(id).value) || 0); document.getElementById('pTotal').value = inputs.reduce((a,b)=>a+b, 0); }

  function hapusPergub(jab) { 
    Swal.fire({
      title: 'Hapus Master Jabatan?',
      text: `Yakin ingin menghapus master jabatan ${jab}?`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#dc3545',
      cancelButtonColor: '#6c757d',
      confirmButtonText: '<i class="bi bi-trash"></i> Ya, Hapus!'
    }).then(async (result) => {
      if (result.isConfirmed) {
        let res = await fetchAPI("hapusPergub", jab); 
        alertSukses(res); 
        muatDataPergub();
      }
    });
  }

  let globalDataPegawai = []; let currentPage = 1; let rowsPerPage = 15; 

  async function muatDataPegawai(forceRefresh = false) {
    if (!forceRefresh && window.cacheDataPegawaiAll && window.cacheDataPegawaiBulan === globalBulanAktif) {
        globalDataPegawai = window.cacheDataPegawaiAll.filter(p => String(p[14] || "PNS").toUpperCase() === globalJenisASN);
        
        let unitList = [...new Set(globalDataPegawai.map(item => item[5]).filter(Boolean))]; 
        let unitDropdown = document.getElementById('filterUnitKerja'); 
        let currentVal = unitDropdown.value;
        unitDropdown.innerHTML = '<option value="">Semua Unit Kerja</option>'; 
        unitList.forEach(s => unitDropdown.innerHTML += `<option value="${s}">${s}</option>`); 
        unitDropdown.value = currentVal;
        
        terapkanFilter();
        return; 
    }

    // Pesan loading yang spesifik agar user tahu sistem tidak error dan sedang bekerja mencari data yang tepat
    startLoading("Memuat Data " + globalJenisASN + " Bulan " + globalBulanAktif + "...");
    
    let data = await fetchAPI("getDaftarPegawai", { bulanAktif: globalBulanAktif, roleUser: currentUser.role, unitkerjaUser: currentUser.unitkerja, jenisASN: "" }); 
    stopLoading();
    
    if (data && !data.error && Array.isArray(data)) { 
        window.cacheDataPegawaiAll = data;
        window.cacheDataPegawaiBulan = globalBulanAktif;
        
        globalDataPegawai = window.cacheDataPegawaiAll.filter(p => String(p[14] || "PNS").toUpperCase() === globalJenisASN);
        
        let unitList = [...new Set(globalDataPegawai.map(item => item[5]).filter(Boolean))]; 
        let unitDropdown = document.getElementById('filterUnitKerja'); 
        unitDropdown.innerHTML = '<option value="">Semua Unit Kerja</option>'; 
        unitList.forEach(s => unitDropdown.innerHTML += `<option value="${s}">${s}</option>`); 
        
        terapkanFilter(); 
    } 
    else { 
        document.getElementById('tabelBodyPegawai').innerHTML = '<tr><td colspan="7" class="text-center text-danger fw-bold">Gagal memuat data dari server.</td></tr>'; 
    }
  }

  function terapkanFilter() {
    let cari = document.getElementById('filterCari').value.toLowerCase(); 
    let fGol = document.getElementById('filterGol').value; 
    let fUnit = document.getElementById('filterUnitKerja').value;
    
    let filtered = globalDataPegawai.filter(row => {
      let nip = String(row[0]).toLowerCase(); 
      let nama = String(row[1]).toLowerCase(); 
      let gol = String(row[3]); 
      let unit = String(row[5]); 
      
      let matchCari = nip.includes(cari) || nama.includes(cari); 
      
      let golDasar = gol.split("/")[0].trim().toUpperCase(); 
      let groupName = "";
      
      if (globalJenisASN === "PPPK") { 
          if(["XVII", "XVI", "XV", "XIV", "XIII"].includes(golDasar)) groupName = "XVII - XIII"; 
          else if(["XII", "XI", "X", "IX"].includes(golDasar)) groupName = "XII - IX"; 
          else if(["VIII", "VII", "VI", "V"].includes(golDasar)) groupName = "VIII - V"; 
          else groupName = "IV - I"; 
      } else { 
          groupName = golDasar; 
      }
      
      let matchGol = (fGol === "" || groupName === fGol); 
      let matchUnit = (fUnit === "" || unit === fUnit); 
      
      return matchCari && matchGol && matchUnit;
    }); 
    
    renderTabelPegawai(filtered);
  }

  function renderTabelPegawai(data) {
    let tbody = document.getElementById('tabelBodyPegawai');
    if(!data.length) { 
        document.getElementById('infoPagination').innerText = "Menampilkan 0 data"; 
        document.getElementById('btnPagination').innerHTML = ""; 
        return tbody.innerHTML = '<tr><td colspan="7" class="text-center text-danger fw-bold">Data tidak ditemukan.</td></tr>'; 
    }
    
    let totalPages = Math.ceil(data.length / rowsPerPage); 
    if(currentPage > totalPages) currentPage = totalPages; 
    let start = (currentPage - 1) * rowsPerPage; 
    let end = start + rowsPerPage; 
    let paginatedData = data.slice(start, end); 
    let html = "";

    paginatedData.forEach(row => {
      let nip = escapeStr(row[0]); let nama = escapeStr(row[1]); let tglLahir = escapeStr(row[2]||""); let gol = escapeStr(row[3]); 
      let unorInduk = escapeStr(row[4]||""); let unit = escapeStr(row[5]); let jab = escapeStr(row[6]); let jenis = escapeStr(row[7]); let stat = escapeStr(row[8]); 
      let gapok = row[9]||0; let tjJab = row[10]||0; let rek = escapeStr(row[11]); let skp = escapeStr(row[12]||""); let uuid = escapeStr(row[15]||""); 
      
      let args = `'${nip}', '${nama}', '${tglLahir}', '${gol}', '${unorInduk}', '${unit}', '${jab}', '${jenis}', '${stat}', ${gapok}, ${tjJab}, '${rek}', '${skp}', '${uuid}'`; 
      let openPanel = `bukaPanel('${nip}', '${nama}', '${gol}', '${jab}', '${stat}', '${unit}')`; 
      
      let btnEdit = globalStatusLock === "Kunci" ? `<button class="btn btn-sm btn-secondary me-1" onclick="alertPeringatan('Periode ini sudah dikunci Admin!')"><i class="bi bi-lock"></i></button>` : `<button class="btn btn-sm btn-primary me-1" onclick="bukaModalPegawai('edit', ${args})">Edit</button>`; 
      let btnDel = globalStatusLock === "Kunci" ? `<button class="btn btn-sm btn-secondary" onclick="alertPeringatan('Periode ini sudah dikunci Admin!')"><i class="bi bi-lock"></i></button>` : `<button class="btn btn-sm btn-danger" onclick="hapusPegawai('${nip}', '${nama}', '${uuid}')">Del</button>`;
      
      html += `<tr class="clickable">
                  <td onclick="${openPanel}">${nip}</td>
                  <td onclick="${openPanel}"><b>${nama}</b></td>
                  <td onclick="${openPanel}">${gol}</td>
                  <td onclick="${openPanel}"><div class="clamp-unit">${unit}</div></td>
                  <td onclick="${openPanel}">${jab}</td>
                  <td onclick="${openPanel}">${rek}</td>
                  <td class="text-center no-click"><button class="btn btn-sm btn-info text-white me-1" onclick="${openPanel}">Buka</button>${btnEdit}${btnDel}</td>
               </tr>`;
    });
    
    tbody.innerHTML = html; 
    document.getElementById('infoPagination').innerText = `Menampilkan ${start + 1} - ${Math.min(end, data.length)} dari Total ${data.length} Pegawai`; 
    
    let btnHtml = `<button class="btn btn-sm btn-outline-primary" onclick="ubahHalaman(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''}>Mundur</button>
                   <span class="btn btn-sm btn-primary disabled text-white fw-bold">Hal ${currentPage} / ${totalPages}</span>
                   <button class="btn btn-sm btn-outline-primary" onclick="ubahHalaman(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''}>Maju</button>`; 
    document.getElementById('btnPagination').innerHTML = btnHtml;
  }

  function ubahHalaman(page) { currentPage = page; terapkanFilter(); }

  function renderDropdownGolongan(selectedValue = "") {
    let sel = document.getElementById('mGolongan'); sel.innerHTML = ""; let html = "";
    if (globalJenisASN === "PNS") { html += `<optgroup label="Golongan IV"><option value="IV/e">IV/e</option><option value="IV/d">IV/d</option><option value="IV/c">IV/c</option><option value="IV/b">IV/b</option><option value="IV/a">IV/a</option></optgroup><optgroup label="Golongan III"><option value="III/d">III/d</option><option value="III/c">III/c</option><option value="III/b">III/b</option><option value="III/a">III/a</option></optgroup><optgroup label="Golongan II"><option value="II/d">II/d</option><option value="II/c">II/c</option><option value="II/b">II/b</option><option value="II/a">II/a</option></optgroup><optgroup label="Golongan I"><option value="I/d">I/d</option><option value="I/c">I/c</option><option value="I/b">I/b</option><option value="I/a">I/a</option></optgroup>`; } 
    else if (globalJenisASN === "PPPK") { html += `<optgroup label="Golongan XVII - XIII"><option value="XVII">XVII</option><option value="XVI">XVI</option><option value="XV">XV</option><option value="XIV">XIV</option><option value="XIII">XIII</option></optgroup><optgroup label="Golongan XII - IX"><option value="XII">XII</option><option value="XI">XI</option><option value="X">X</option><option value="IX">IX</option></optgroup><optgroup label="Golongan VIII - V"><option value="VIII">VIII</option><option value="VII">VII</option><option value="VI">VI</option><option value="V">V</option></optgroup><optgroup label="Golongan IV - I"><option value="IV">IV</option><option value="III">III</option><option value="II">II</option><option value="I">I</option></optgroup>`; }
    sel.innerHTML = html; if (selectedValue) { sel.value = selectedValue; }
  }

  function bukaModalPegawai(aksi, nip="", nama="", tglLahir="", golongan="", unorInduk="", unitkerja="", jabatan="", jenis="Struktural", status="TK/0 = 1", gapok=0, tjJab=0, rekening="", skp="", uuid="") {
    document.getElementById('mAksi').value = aksi; document.getElementById('mNipLama').value = nip; document.getElementById('mTglLahir').value = tglLahir; document.getElementById('mGolongan').value = golongan; document.getElementById('mUuid').value = uuid; renderDropdownGolongan(golongan); document.getElementById('mJenis').value = jenis; 
    
    document.getElementById('mStatus').value = status; 

  document.getElementById('mGapok').value = formatRupiah(gapok); 
  document.getElementById('mTjJab').value = formatRupiah(tjJab); 
    
    document.getElementById('mRekening').value = rekening;
    document.getElementById('mSkp').value = skp || "";
    
    let elNip = document.getElementById('mNip'); let elNama = document.getElementById('mNama'); let elUnitKerja = document.getElementById('mUnitKerja'); let elUnorInduk = document.getElementById('mUnorInduk');
    elNip.value = nip; elNama.value = nama; elUnitKerja.value = unitkerja;
    
    let opdAktif = document.getElementById('loginNamaDinas').innerText;
    elUnorInduk.value = unorInduk || opdAktif;
    elUnitKerja.value = unitkerja || opdAktif;

    let isOperatorEdit = (aksi === 'edit' && currentUser.role === "Operator");
    if (isOperatorEdit) { elNip.readOnly = true; elNip.classList.add("readonly-field"); elNama.readOnly = true; elNama.classList.add("readonly-field"); } 
    else { elNip.readOnly = false; elNip.classList.remove("readonly-field"); elNama.readOnly = false; elNama.classList.remove("readonly-field"); }
    
    if (currentUser.role === "Operator") { 
        elUnitKerja.value = currentUser.unitkerja;  
        elUnitKerja.readOnly = true; elUnitKerja.classList.add("readonly-field"); 
    } 
    else if (currentUser.role === "Admin Sub Unit/UPTD") {
        elUnitKerja.readOnly = false; elUnitKerja.classList.remove("readonly-field");
        let unitsInSubUnit = arrayUnitKerjaFull.filter(u => String(u.subUnit).toLowerCase() === String(currentUser.unitkerja).toLowerCase()).map(u => u.nama);
        setupAutocomplete(elUnitKerja, unitsInSubUnit);
    }
    else { 
        elUnitKerja.readOnly = false; elUnitKerja.classList.remove("readonly-field"); 
        setupAutocomplete(elUnitKerja, arrayUnitKerjaValid);
    }
    
    muatDropdownJabatan(jabatan); document.getElementById('pegawaiModalTitle').innerText = aksi === 'edit' ? 'Edit Pegawai' : 'Tambah Pegawai Baru'; 
    let modalObj = bootstrap.Modal.getInstance(document.getElementById('modalKelolaPegawai')) || new bootstrap.Modal(document.getElementById('modalKelolaPegawai'));
    modalObj.show();
  }

  async function submitPegawai(e) {
    e.preventDefault(); 
    let aksi = document.getElementById('mAksi').value; 
    let nipCek = document.getElementById('mNip').value.trim(); 
    if (!/^\d{18}$/.test(nipCek)) { return alertPeringatan("Gagal menyimpan! NIP belum valid (Wajib 18 digit angka)."); } 
    
    let s = document.getElementById('mUnitKerja').value.trim(); 
    if(currentUser.role === 'Admin' && !arrayUnitKerjaValid.includes(s)) { return alertPeringatan("Unit Kerja tidak valid!"); } 
    
    let j = document.getElementById('mJabatan').value.trim(); 
    if(!arrayJabatanValid.includes(j)) { return alertPeringatan("Jabatan tidak valid!"); }
    
    startLoading("Menyimpan Pegawai..."); 
    
    const d = { 
      nipLama: document.getElementById('mNipLama').value, 
      nip: nipCek, 
      nama: document.getElementById('mNama').value, 
      tglLahir: document.getElementById('mTglLahir').value, 
      golongan: document.getElementById('mGolongan').value, 
      unorInduk: document.getElementById('mUnorInduk').value, 
      unitkerja: s, 
      namaJabatan: j, 
      jenisJab: document.getElementById('mJenis').value, 
      statusKawin: document.getElementById('mStatus').value, 
      
      gapok: unformatRupiah(document.getElementById('mGapok').value), 
      tjJab: unformatRupiah(document.getElementById('mTjJab').value),
      rekening: document.getElementById('mRekening').value, 
      skp: document.getElementById('mSkp').value, 
      bulan: globalBulanAktif, 
      statusPegawai: globalJenisASN, 
      // 👇 INI YANG BIKIN ERROR KEMARIN, SUDAH SAYA BERSIHKAN!
      uuid: document.getElementById('mUuid').value || "", 
      roleUser: currentUser.role, 
      unitkerjaUser: currentUser.unitkerja  
    };

    let res = await fetchAPI(aksi === 'edit' ? "updatePegawai" : "simpanPegawai", d); 
    stopLoading(); 
    
    // 👇 BACA RESPON LEBIH CERDAS DAN KEBAL ERROR
    let isError = false;
    let pesanInfo = "";

    if (typeof res === "object" && res !== null) {
        isError = (res.status === "error");
        pesanInfo = res.pesan || "Tidak ada pesan dari server";
    } else {
        let strRes = String(res).toLowerCase();
        isError = strRes.includes("gagal") || strRes.includes("error");
        pesanInfo = String(res);
    }

    if (isError) { 
        alertError(pesanInfo); 
    } else { 
        alertSukses(pesanInfo);
        if(window.cacheDetailPegawai) {
      delete window.cacheDetailPegawai[d.nip + "_" + d.bulan];
      if (aksi === 'edit') delete window.cacheDetailPegawai[d.nipLama + "_" + d.bulan];
  } 
        let modalObj = bootstrap.Modal.getInstance(document.getElementById('modalKelolaPegawai'));
        if (modalObj) modalObj.hide(); 
        document.getElementById('formPegawai').reset(); 
        
        // Tarik ulang data fresh dari database Master OPD
        muatDataPegawai(true); 
    }
  }

  function hapusPegawai(nip, nama, uuid) { 
    Swal.fire({
      title: 'Hapus Pegawai?',
      text: `Yakin ingin menghapus ${nama} dari bulan ini saja?`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#dc3545',
      cancelButtonColor: '#6c757d',
      confirmButtonText: '<i class="bi bi-trash"></i> Ya, Hapus!'
    }).then(async (result) => {
      if (result.isConfirmed) {
        startLoading("Menghapus Pegawai..."); 
        let res = await fetchAPI("hapusPegawai", {nip: nip, bulan: globalBulanAktif, uuid: uuid}); 
        stopLoading(); 
        
        let pesan = String(res).toLowerCase(); 
        if (pesan.includes("gagal") || pesan.includes("error")) { 
            alertError(res.pesan || res); 
        } else { 
            alertSukses(res); 
            
            if(window.cacheDataPegawaiAll) {
                window.cacheDataPegawaiAll = window.cacheDataPegawaiAll.filter(p => !(p[0] === nip && p[13].toLowerCase() === globalBulanAktif.toLowerCase()));
            }
            globalDataPegawai = window.cacheDataPegawaiAll.filter(p => String(p[14] || "PNS").toUpperCase() === globalJenisASN);
            terapkanFilter(); 
        }
      }
    });
  }

  // Paste kode ini di js_main.js Anda

// 1. Kunci keyboard saat ngetik (Mentok 18 angka & nolak huruf)
// 1. Kunci keyboard saat ngetik & Ubah input uang ke TEXT
document.addEventListener('DOMContentLoaded', function() {
    let inputNip = document.getElementById('mNip');
    if(inputNip) {
        inputNip.addEventListener('input', function(e) {
            this.value = this.value.replace(/[^0-9]/g, ''); 
            if(this.value.length > 18) this.value = this.value.slice(0, 18); 
        });
    }

    // Ubah otomatis input angka menjadi text agar bisa pakai titik ribuan
    let inputsUang = ['mGapok', 'mTjJab', 'aGapok', 'aTjJab', 'aBulat', 'aTjIstri', 'aTjAnak', 'aJmlKeluarga', 'aTjBeras', 'aTjPajak', 'aTjBPJS4', 'aTjJKK', 'aTjJKM', 'aJmlKotor', 'aPotIWP8', 'aPotIWP1', 'aPotBPJS4', 'aPotPajak', 'aPotJKK', 'aPotJKM', 'aJmlPotongan', 'aJmlBersih'];
    
    inputsUang.forEach(id => {
        let el = document.getElementById(id);
        if(el) {
            el.type = 'text'; // Paksa HTML ganti jadi Text
            if(id === 'mGapok' || id === 'mTjJab' || id === 'aTjPajak' || id === 'aPotPajak') {
                el.addEventListener('input', function() {
                    this.value = formatRupiah(this.value); // Pasang titik real-time
                });
            }
        }
    });
});

// 2. Peringatan saat user pindah kolom (onblur)
function validasiNIP(input) { 
    let nilaiBersih = input.value.replace(/[^0-9]/g, ''); 
    input.value = nilaiBersih; 
    
    if (nilaiBersih !== "") { 
        if (nilaiBersih.length !== 18) { 
            // Jika kurang dari 18 digit
            alertPeringatan("Perhatian! NIP wajib terdiri dari TEPAT 18 digit angka murni. Saat ini baru " + nilaiBersih.length + " digit."); 
            input.classList.add('border-danger', 'is-invalid'); 
            input.classList.remove('border-primary');
        } else { 
            // Jika pas 18 digit
            input.classList.remove('border-danger', 'is-invalid'); 
            input.classList.add('border-success', 'is-valid');
        } 
    } 
}

  async function inisialisasiAplikasi() {
    let bAktif = sessionStorage.getItem('globalBulanAktif') || "";
    let jASN = sessionStorage.getItem('globalJenisASN') || "";
    let cUser = sessionStorage.getItem('isLoggedIn') === 'true' ? JSON.parse(sessionStorage.getItem('currentUser')) : null;

    let payloadData = {};
    if (cUser && bAktif) {
        payloadData = { bulanAktif: bAktif, roleUser: cUser.role, unitkerjaUser: cUser.unitkerja, jenisASN: "" };
    }

    let cachedDataStr = localStorage.getItem('simTppCacheData');
    let isCached = false;
    
    if (cachedDataStr) {
        try {
            let cachedData = JSON.parse(cachedDataStr);
            terapkanDataInit(cachedData, cUser, bAktif, jASN, false);
            isCached = true;
        } catch(e) { 
            console.error("Cache Error:", e);
            isCached = false; 
        }
    }

    if (!isCached) {
        startLoading("Memuat Data Sistem...");
    }

    fetchAPI("getInitialLoad", payloadData).then(dataInit => {
        if (!isCached) stopLoading(); 
        
        if (dataInit && !dataInit.error) {
            localStorage.setItem('simTppCacheData', JSON.stringify(dataInit));
            terapkanDataInit(dataInit, cUser, bAktif, jASN, isCached); 
        } else if (dataInit.error || dataInit.status === "error") {
            if (!isCached) alertError(dataInit.pesan || "Gagal memuat data dari server.");
        }
    }).catch(err => {
        console.error("Fetch Error:", err);
        if (!isCached) {
            stopLoading();
            alertError("Terjadi kesalahan sistem: " + err.message);
        }
    });
  }

  async function bukaPanel(nip, nama, gol, jabatan, statusTER, unitKerja) {
    asNIPAktif = nip; 
    statusTERAktif = statusTER || ""; 
    objNominatifSetahun = null;
    
    let objPer = arrayPeriode.find(x => x.namaPeriode === globalBulanAktif); 
 
    globalJenisPeriode = objPer ? (objPer.jenisPeriode || "Reguler") : "Reguler";
    
    document.getElementById('labelNamaASN').innerText = nama; 
    document.getElementById('labelStatusASN').innerText = "Status ASN: " + globalJenisASN; 
    document.getElementById('labelNIPASN').innerText = "NIP: " + nip; 
    document.getElementById('labelGolonganASN').innerText = "Gol: " + gol; 
    document.getElementById('labelJabatanASN').innerText = "Jabatan: " + jabatan;
    
    if(document.getElementById('labelBulanASNPanel')) document.getElementById('labelBulanASNPanel').innerText = globalBulanAktif; 
    if(document.getElementById('labelBulanGajiBersih')) document.getElementById('labelBulanGajiBersih').innerText = "(" + globalBulanAktif + ")"; 
    if(document.getElementById('labelBulanTerBawah')) document.getElementById('labelBulanTerBawah').innerText = globalBulanAktif; 
    if(document.getElementById('labelBulanNominatif')) document.getElementById('labelBulanNominatif').innerText = globalBulanAktif;
    
    let elUnit = document.getElementById('labelUnitKerjaASN'); 
    if(elUnit) elUnit.innerText = "Unit: " + (unitKerja || "");
    
    //  TIDAK DI PAKAI ---document.getElementById('formGaji').reset(); 
    document.getElementById('formAbsen').reset(); 
    document.getElementById('inpBulan').value = globalBulanAktif;  
    document.getElementById('inpHariKerja').value = globalHariKerja; 
    
    // HAPUS CLASS HIDDEN AGAR ISINYA TERLIHAT
    document.getElementById('viewManajemenASN').classList.remove('hidden');

    // MUNCULKAN PANEL SEBAGAI POPUP MODAL RAKSASA
    let myModal = bootstrap.Modal.getInstance(document.getElementById('modalPopupASN')) || new bootstrap.Modal(document.getElementById('modalPopupASN'));
    myModal.show();
    
    // 👇 IMPLEMENTASI CACHE PINTAR 👇
    window.cacheDetailPegawai = window.cacheDetailPegawai || {};

    let cacheKey = nip + "_" + globalBulanAktif;
    let res;

    if (window.cacheDetailPegawai[cacheKey]) {
        // Data sudah ada di memori! Tarik instan tanpa loading ke server
        res = window.cacheDetailPegawai[cacheKey];
    } else {
        // Data belum ada, tarik dari server (Loading 5 detik)
        startLoading("Menarik Riwayat Gaji & Absen...");
        await new Promise(resolve => setTimeout(resolve, 50));
        
        res = await fetchAPI("getDetailASN", {nip: nip, bulan: globalBulanAktif});
        stopLoading();
        
        if (res && !res.error) {
            window.cacheDetailPegawai[cacheKey] = res; // Simpan ke memori agar besok gak loading lagi
        }
    }

    // --- Sisa kodenya sama persis ---
    if(res && !res.error) { 
      isGajiTersimpan = (res.gaji) ? true : false; 
      isAbsenTersimpan = (res.absen) ? true : false; 
      navigasiTab('tabGaji', true); 
      
      if(res.pergub) { baseTPP = res.pergub; } 
      
      if(res.pegawaiInfo) { 
        document.getElementById('aGapok').value = formatRupiah(res.pegawaiInfo.gapok || 0); 
        document.getElementById('aTjJab').value = formatRupiah(res.pegawaiInfo.tjJab || 0); 
        
        if(res.pegawaiInfo.skpBulanan && res.pegawaiInfo.skpBulanan !== "") {
          document.getElementById('inpSKP').value = res.pegawaiInfo.skpBulanan;
        } else {
          document.getElementById('inpSKP').value = "Baik"; 
        }
      } 
        
      const setV = (id, val) => { 
        let el = document.getElementById(id); 
        if(el) el.value = (val !== undefined && val !== null && !isNaN(val)) ? parseInt(val) : 0; 
      }; 
      
      if(res.gaji) { 
        if(document.getElementById('aTjPajak')) document.getElementById('aTjPajak').value = formatRupiah(res.gaji[10]); 
        if(document.getElementById('aPotPajak')) document.getElementById('aPotPajak').value = formatRupiah(res.gaji[17]); 
      } else { 
        if(document.getElementById('aTjPajak')) document.getElementById('aTjPajak').value = 0; 
        if(document.getElementById('aPotPajak')) document.getElementById('aPotPajak').value = 0; 
      } 
      
      hitungAmprah(); 
      
      // 👇 PERUBAHAN 4: Simpan Gaji Diam-Diam di Belakang Layar 👇
      if (globalStatusLock !== "Kunci") {
          simpanGajiSiluman(); 
      } else {
          isGajiTersimpan = true; 
      }
      
      if(res.absen) { 
        let hkDb = parseInt(res.absen[3]); 
        if(hkDb === globalHariKerja6) { document.getElementById('inpPolaHK').value = "6"; } 
        else { document.getElementById('inpPolaHK').value = "5"; } 
      } else { 
        document.getElementById('inpPolaHK').value = "5"; 
      } 
      ubahPolaHK(); 
      
      let btnGaji = document.getElementById('btnSimpanGaji'); 
      let btnTPP = document.getElementById('btnSimpanTPP'); 
      
      if (globalStatusLock === "Kunci") { 
        if(btnGaji) { btnGaji.disabled = true; btnGaji.innerHTML = "<i class='bi bi-lock'></i> TERKUNCI"; btnGaji.classList.replace('btn-primary', 'btn-secondary'); }
        if(btnTPP) { btnTPP.disabled = true; btnTPP.innerHTML = "<i class='bi bi-lock'></i> TERKUNCI"; btnTPP.classList.replace('btn-primary', 'btn-secondary'); }
      } else { 
        if(btnGaji) { btnGaji.disabled = false; btnGaji.innerHTML = "<i class='bi bi-save'></i> SIMPAN AMPRAH GAJI"; btnGaji.classList.replace('btn-secondary', 'btn-primary'); }
        if(btnTPP) { btnTPP.disabled = false; btnTPP.innerHTML = "<i class='bi bi-save'></i> SIMPAN PERHITUNGAN TPP"; btnTPP.classList.replace('btn-secondary', 'btn-primary'); }
      }
      
      if(res.absen) { 
        let skpRaw = String(res.absen[4] || "Baik|Menilai"); 
        let skpParts = skpRaw.split("|"); 
        
        if (!res.pegawaiInfo || !res.pegawaiInfo.skpBulanan || res.pegawaiInfo.skpBulanan === "") {
             document.getElementById('inpSKP').value = skpParts[0] || "Baik"; 
        }
        
        document.getElementById('chkTidakMenilai').checked = (skpParts[1] === "Tidak Menilai"); 
        setV('vDL', res.absen[5]); setV('vS', res.absen[6]); setV('vC', res.absen[7]); setV('vKP', res.absen[8]); setV('vTL1', res.absen[9]); setV('vTL2', res.absen[10]); setV('vTL3', res.absen[11]); setV('vTL4', res.absen[12]); setV('vCP1', res.absen[13]); setV('vCP2', res.absen[14]); setV('vCP3', res.absen[15]); setV('vCP4', res.absen[16]); setV('vTK', res.absen[17]); setV('vASUB', res.absen[18]); 
      } else { 
        document.getElementById('chkTidakMenilai').checked = false; 
        ['vDL','vS','vC','vKP','vTL1','vTL2','vTL3','vTL4','vCP1','vCP2','vCP3','vCP4','vTK','vASUB'].forEach(id => { if(document.getElementById(id)) document.getElementById(id).value = 0; }); 
      } 
      
      hitungMatriksTPP();
    }
  }

  function hitungAmprah(isManualPajak = false) {
    // 1. Baca nilai dan hapus titiknya dulu
    let gapok = unformatRupiah(document.getElementById('aGapok')?.value); 
    let tjJab = unformatRupiah(document.getElementById('aTjJab')?.value); 
    let statTer = statusTERAktif || "TK/0 = 1"; 
    let isKawin = statTer.startsWith("K"); 
    let jmlJiwa = parseInt(statTer.split("=")[1]) || 1; 
    let jmlAnak = jmlJiwa - 1 - (isKawin ? 1 : 0); 
    if(jmlAnak < 0) jmlAnak = 0;
    
    let tjIstri = Math.round(isKawin ? (gapok * 0.10) : 0); 
    let tjAnak = Math.round(gapok * 0.02 * jmlAnak); 
    let jmlKeluarga = gapok + tjIstri + tjAnak; 
    let tjBeras = (gapok > 0) ? Math.round(72420 * jmlJiwa) : 0; 
    
    let basisBPJS = jmlKeluarga + tjJab; 
    let tjBPJS4 = Math.round(basisBPJS * 0.04); 
    let tjJKK = Math.round(gapok * 0.0024); 
    let tjJKM = Math.round(gapok * 0.0072); 
    
    if (globalJenisPeriode !== "Reguler") { tjBPJS4 = 0; tjJKK = 0; tjJKM = 0; }
    
    if (!isManualPajak) {
      let dppTER = gapok + tjIstri + tjAnak + tjJab + tjBeras + tjBPJS4 + tjJKK + tjJKM; 
      let statusKawinShort = (isKawin ? "K/" : "TK/") + jmlAnak; 
      let katTER = "A"; 
      if(["TK/0", "TK/1", "K/0"].includes(statusKawinShort)) katTER = "A"; 
      else if(["TK/2", "TK/3", "K/1", "K/2"].includes(statusKawinShort)) katTER = "B"; 
      else if(["K/3"].includes(statusKawinShort)) katTER = "C";
      
      const terA = [[5400000,0],[5650000,0.25],[5950000,0.5],[6300000,0.75],[6750000,1],[7500000,1.25],[8550000,1.5],[9650000,1.75],[10050000,2],[10350000,2.25],[10700000,2.5],[11050000,3],[11600000,3.5],[12500000,4],[13750000,5],[15100000,6],[16950000,7],[19750000,8],[24100000,9],[26450000,10],[28000000,11],[30000000,12],[32000000,13],[35000000,14],[37000000,15],[39000000,16],[41000000,17],[43000000,18],[46000000,19],[49000000,20],[52000000,21],[55000000,22],[58000000,23],[61000000,24],[64000000,25],[67000000,26],[70000000,27],[74000000,28],[78000000,29],[82000000,30],[86000000,31],[90000000,32],[94000000,33],[Infinity,34]]; 
      const terB = [[6200000,0],[6500000,0.25],[6850000,0.5],[7300000,0.75],[9200000,1],[10750000,1.5],[11250000,2],[11600000,2.5],[12600000,3],[13600000,4],[14950000,5],[16400000,6],[18450000,7],[21850000,8],[26000000,9],[27700000,10],[29350000,11],[31450000,12],[33450000,13],[37000000,14],[39000000,15],[41000000,16],[43000000,17],[46000000,18],[48000000,19],[51000000,20],[54000000,21],[57000000,22],[60000000,23],[63000000,24],[66000000,25],[69000000,26],[73000000,27],[77000000,28],[81000000,29],[85000000,30],[89000000,31],[93000000,32],[97000000,33],[Infinity,34]]; 
      const terC = [[6600000,0],[6950000,0.25],[7350000,0.5],[7800000,0.75],[8850000,1],[9800000,1.25],[10950000,1.5],[11200000,2],[11600000,2.5],[12050000,3],[13200000,4],[14400000,5],[15900000,6],[17950000,7],[21200000,8],[25850000,9],[27200000,10],[28900000,11],[30800000,12],[32800000,13],[35800000,14],[38000000,15],[40000000,16],[42000000,17],[44000000,18],[47000000,19],[50000000,20],[53000000,21],[56000000,22],[59000000,23],[62000000,24],[65000000,25],[68000000,26],[71000000,27],[75000000,28],[79000000,29],[83000000,30],[87000000,31],[91000000,32],[95000000,33],[Infinity,34]];
      
      let tabelPilihan = katTER === "A" ? terA : (katTER === "B" ? terB : terC); 
      let pctTER = 0; 
      for(let i=0; i<tabelPilihan.length; i++) { 
        if(dppTER <= tabelPilihan[i][0]) { pctTER = tabelPilihan[i][1]; break; } 
      }
      let tjPajakOtomatis = Math.round(dppTER * (pctTER / 100)); 
      if(document.getElementById('aTjPajak')) { 
        document.getElementById('aTjPajak').value = formatRupiah(tjPajakOtomatis); 
        document.getElementById('aPotPajak').value = formatRupiah(tjPajakOtomatis); 
      }
    }
    
    let tjPajak = unformatRupiah(document.getElementById('aTjPajak')?.value); 
    let potPajak = unformatRupiah(document.getElementById('aPotPajak')?.value); 
    let potIWP8 = Math.round(jmlKeluarga * 0.08); 
    let potIWP1 = Math.round(basisBPJS * 0.01); 
    let potJKK = tjJKK; 
    let potJKM = tjJKM; 
    
    if (globalJenisPeriode !== "Reguler") { potIWP8 = 0; potIWP1 = 0; potJKK = 0; potJKM = 0; }
    
    let totalPotongan = potIWP8 + potIWP1 + tjBPJS4 + potJKK + potJKM + potPajak; 
    let kotorSementara = gapok + tjIstri + tjAnak + tjJab + tjBeras + tjBPJS4 + tjJKK + tjJKM + tjPajak; 
    let bersihSementara = kotorSementara - totalPotongan; 
    let sisa = bersihSementara % 100; 
    let pembulatan = (sisa > 0) ? (100 - sisa) : 0; 
    let jmlKotor = kotorSementara + pembulatan; 
    let gajiBersih = jmlKotor - totalPotongan;
    
    // 2. Kembalikan nilainya ke HTML dengan format Titik Ribuan
    if(document.getElementById('aJmlKeluarga')) { 
      document.getElementById('aJmlKeluarga').value = formatRupiah(jmlKeluarga); 
      document.getElementById('aTjIstri').value = formatRupiah(tjIstri); 
      document.getElementById('aTjAnak').value = formatRupiah(tjAnak); 
      document.getElementById('aTjBeras').value = formatRupiah(tjBeras); 
      document.getElementById('aTjBPJS4').value = formatRupiah(tjBPJS4); 
      document.getElementById('aTjJKK').value = formatRupiah(tjJKK); 
      document.getElementById('aTjJKM').value = formatRupiah(tjJKM); 
      document.getElementById('aPotJKK').value = formatRupiah(potJKK); 
      document.getElementById('aPotJKM').value = formatRupiah(potJKM); 
      document.getElementById('aPotBPJS4').value = formatRupiah(tjBPJS4); 
      document.getElementById('aPotIWP1').value = formatRupiah(potIWP1); 
      document.getElementById('aPotIWP8').value = formatRupiah(potIWP8); 
      document.getElementById('aBulat').value = formatRupiah(pembulatan); 
      document.getElementById('aJmlKotor').value = formatRupiah(jmlKotor); 
      document.getElementById('aJmlPotongan').value = formatRupiah(totalPotongan); 
      document.getElementById('aJmlBersih').value = formatRupiah(gajiBersih); 
    }
  }

  function hitungMatriksTPP() {
    if(!baseTPP || typeof baseTPP.bk === 'undefined') return; 
    const fR = (val) => Math.round(val).toLocaleString('id-ID'); 
    const getV = (id) => parseFloat(document.getElementById(id).value) || 0;
    
    let hk = parseInt(document.getElementById('inpHariKerja').value) || 0; 
    let totalTidakMasuk = getV('vDL') + getV('vS') + getV('vC') + getV('vKP') + getV('vTK'); 
    let totalMasuk = hk - totalTidakMasuk; 
    if(totalMasuk < 0) totalMasuk = 0; 
    document.getElementById('txtHariTidakMasuk').innerText = totalTidakMasuk + " Hari"; 
    document.getElementById('txtHariMasuk').innerText = totalMasuk + " Hari";
    
    let pctAbsPK = (getV('vS')*0.03) + (getV('vC')*0.03) + (getV('vTK')*0.03) + (getV('vTL1')*0.005) + (getV('vTL2')*0.01) + (getV('vTL3')*0.0125) + (getV('vTL4')*0.015) + (getV('vCP1')*0.005) + (getV('vCP2')*0.01) + (getV('vCP3')*0.0125) + (getV('vCP4')*0.0155);
    if(pctAbsPK > 1) pctAbsPK = 1; 
    
    let asubV = getV('vASUB'); 
    let pctASUB = (asubV < 28) ? (asubV * 0.02) : 1.0; 
    if (asubV === 0) pctASUB = 0; 
    let pctAbsDK = pctAbsPK + pctASUB; 
    if(pctAbsDK > 1) pctAbsDK = 1;
    
    let skp = document.getElementById('inpSKP').value; 
    let isChecked = document.getElementById('chkTidakMenilai').checked; 
    let bobotSKP = 100; 
    if(skp === "Sangat Baik" || skp === "Baik") bobotSKP = 100; 
    else if(skp === "Cukup") bobotSKP = 90; 
    else if(skp === "Kurang") bobotSKP = 75; 
    else if(skp === "Sangat Kurang") bobotSKP = 50; 
    else if(skp === "Tidak Buat") bobotSKP = 40; 
    
    if (isChecked) bobotSKP -= 5; 
    let rasioSKP = bobotSKP / 100; 
    
    document.getElementById('txtSkp').innerText = bobotSKP + "%"; 
    document.getElementById('txtAbsPk').innerText = (pctAbsPK*100).toFixed(2) + "%"; 
    document.getElementById('txtAbsDk').innerText = (pctAbsDK*100).toFixed(2) + "%";
    
    const keys = ['bk', 'pk', 'kk', 'tb', 'kp', 'total'];
    keys.forEach(k => { 
      let kotor = baseTPP[k] || 0; 
      let bPK = kotor * 0.60; 
      let nSKP = bPK * rasioSKP; 
      let pAbsPK = nSKP * pctAbsPK; 
      let akhPK = nSKP - pAbsPK; 
      let bDK = kotor * 0.40; 
      let pAbsDK = bDK * pctAbsDK; 
      let akhDK = bDK - pAbsDK; 
      let totKotor = kotor; 
      let totPot = pAbsPK + pAbsDK; 
      let hasilAkhir = akhPK + akhDK; 
      
      if(document.getElementById('mKotor_'+k)) document.getElementById('mKotor_'+k).innerText = fR(kotor); 
      if(document.getElementById('mBpk_'+k)) document.getElementById('mBpk_'+k).innerText = fR(bPK); 
      if(document.getElementById('mNskp_'+k)) document.getElementById('mNskp_'+k).innerText = fR(nSKP); 
      if(document.getElementById('mPabspk_'+k)) document.getElementById('mPabspk_'+k).innerText = fR(pAbsPK); 
      if(document.getElementById('mAkhpk_'+k)) document.getElementById('mAkhpk_'+k).innerText = fR(akhPK); 
      if(document.getElementById('mBdk_'+k)) document.getElementById('mBdk_'+k).innerText = fR(bDK); 
      if(document.getElementById('mPabsdk_'+k)) document.getElementById('mPabsdk_'+k).innerText = fR(pAbsDK); 
      if(document.getElementById('mAkhdk_'+k)) document.getElementById('mAkhdk_'+k).innerText = fR(akhDK); 
      if(document.getElementById('mTotKotor_'+k)) document.getElementById('mTotKotor_'+k).innerText = fR(totKotor); 
      if(document.getElementById('mTotPot_'+k)) document.getElementById('mTotPot_'+k).innerText = fR(totPot); 
      if(document.getElementById('mHasil_'+k)) document.getElementById('mHasil_'+k).innerText = fR(hasilAkhir); 
    });
  }


  function ubahPolaHK() { let pola = document.getElementById('inpPolaHK').value; document.getElementById('inpHariKerja').value = (pola === "5") ? globalHariKerja : globalHariKerja6; hitungMatriksTPP(); }

  // Fungsi baru pengganti submitGaji
  async function simpanGajiSiluman() {
      const d = { 
          nip: asNIPAktif, bulan: globalBulanAktif, 
          gapok: unformatRupiah(document.getElementById('aGapok').value), 
          tjIstri: unformatRupiah(document.getElementById('aTjIstri').value), 
          tjAnak: unformatRupiah(document.getElementById('aTjAnak').value), 
          jumlahKeluarga: unformatRupiah(document.getElementById('aJmlKeluarga').value), 
          tjJabatan: unformatRupiah(document.getElementById('aTjJab').value), 
          tjTerpencil: 0, tkd: 0, 
          tjBeras: unformatRupiah(document.getElementById('aTjBeras').value), 
          tjPajak: unformatRupiah(document.getElementById('aTjPajak').value), 
          tjBPJS4: unformatRupiah(document.getElementById('aTjBPJS4').value), 
          tjJKK: unformatRupiah(document.getElementById('aTjJKK').value), 
          tjJKM: unformatRupiah(document.getElementById('aTjJKM').value), 
          taperaPK: 0, 
          pembulatan: unformatRupiah(document.getElementById('aBulat').value), 
          jmlKotor: unformatRupiah(document.getElementById('aJmlKotor').value), 
          potPajak: unformatRupiah(document.getElementById('aPotPajak').value), 
          potBPJS: unformatRupiah(document.getElementById('aPotBPJS4').value), 
          potIWP1: unformatRupiah(document.getElementById('aPotIWP1').value), 
          potIWP8: unformatRupiah(document.getElementById('aPotIWP8').value), 
          potTaperum: 0, 
          potJKK: unformatRupiah(document.getElementById('aPotJKK').value), 
          potJKM: unformatRupiah(document.getElementById('aPotJKM').value), 
          jmlPotongan: unformatRupiah(document.getElementById('aJmlPotongan').value), 
          jmlBersih: unformatRupiah(document.getElementById('aJmlBersih').value) 
      };

      // Tembak server secara async tanpa memblokir layar (Siluman)
      fetchAPI("simpanGajiOtomatis", d).then(res => {
          if (res && res.status !== "error") {
              isGajiTersimpan = true;
              if(window.cacheDetailPegawai) delete window.cacheDetailPegawai[asNIPAktif + "_" + globalBulanAktif];
          }
      }).catch(e => console.log("Gagal simpan gaji otomatis", e));
  }


  async function submitPerhitunganTPP(e) {
    e.preventDefault(); const getV = (id) => document.getElementById(id).value; startLoading("Menyimpan Kehadiran..."); let statusMenilai = document.getElementById('chkTidakMenilai').checked ? "Tidak Menilai" : "Menilai"; let gabunganSKP = getV('inpSKP') + "|" + statusMenilai; let hkDB = document.getElementById('inpHariKerja').value; const d = { nip: asNIPAktif, bulan: globalBulanAktif, hariKerja: hkDB, skp: gabunganSKP, dl: getV('vDL'), s: getV('vS'), c: getV('vC'), kp: getV('vKP'), tk: getV('vTK'), asub: getV('vASUB'), tl1: getV('vTL1'), tl2: getV('vTL2'), tl3: getV('vTL3'), tl4: getV('vTL4'), cp1: getV('vCP1'), cp2: getV('vCP2'), cp3: getV('vCP3'), cp4: getV('vCP4') };
    
    let res = await fetchAPI("simpanPerhitunganTPP", d); 
  stopLoading(); 
  
  // PERBAIKAN DI SINI
  if(res && res.status === "error") { 
      alertError(res.pesan); 
  } else { 
      alertSukses(res.pesan);
      if(window.cacheDetailPegawai) delete window.cacheDetailPegawai[asNIPAktif + "_" + globalBulanAktif];
      isAbsenTersimpan = true; 
      hitungPajakTahunan(); 
      muatNominatif(); 
  }
  }

    
 async function hitungPajakTahunan() {
    document.getElementById('loaderTahunan').classList.remove('hidden');
    document.getElementById('hasilTahunan').classList.add('hidden');

    const getV = (id) => document.getElementById(id).value;
    let statusMenilai = document.getElementById('chkTidakMenilai').checked ? "Tidak Menilai" : "Menilai";
    let gabunganSKP = getV('inpSKP') + "|" + statusMenilai;
    let customAbsen = { 
        skp: gabunganSKP, s: getV('vS'), c: getV('vC'), tk: getV('vTK'), 
        tl1: getV('vTL1'), tl2: getV('vTL2'), tl3: getV('vTL3'), tl4: getV('vTL4'), 
        cp1: getV('vCP1'), cp2: getV('vCP2'), cp3: getV('vCP3'), cp4: getV('vCP4'), asub: getV('vASUB') 
    };

    let res = await fetchAPI("hitungNominatif", {
        nip: asNIPAktif,
        bulanAktif: globalBulanAktif,
        customAbsen: customAbsen,
        refBulanGaji: globalRefBulanGaji
    });

    document.getElementById('loaderTahunan').classList.add('hidden');
    
    // 👇 CEGAT ERROR DARI SERVER AGAR LAYAR TIDAK ZONK / NaN 👇
    if (res.status === "error") return alertError(res.pesan);
    if (res.error) return alertPeringatan("Peringatan: " + res.error);

    objNominatifSetahun = res;
    document.getElementById('hasilTahunan').classList.remove('hidden');
    
    const fRp = (angka) => "Rp " + Math.round(angka || 0).toLocaleString('id-ID');

    if(document.getElementById('labelBulanPencairan')) document.getElementById('labelBulanPencairan').innerText = globalRefBulanGaji || globalBulanAktif;
    if(document.getElementById('labelBulanPencairan2')) document.getElementById('labelBulanPencairan2').innerText = globalRefBulanGaji || globalBulanAktif;

    // 👇 LOGIKA LAYOUT DESEMBER vs TER 👇
    if (res.isModeDesember === true && res.akumulasi) {
        
        // BUKA DESEMBER, TUTUP TER
        if(document.getElementById('layoutTER')) document.getElementById('layoutTER').classList.add('hidden');
        if(document.getElementById('layoutDesember')) document.getElementById('layoutDesember').classList.remove('hidden');

        let tbodyRiwayat = document.getElementById('tabelRiwayatDesember');
        tbodyRiwayat.innerHTML = "";

        // Variabel penampung subtotal Jan-Nov
        let subGaji = 0, subTpp = 0, subIwp = 0, subPphGaji = 0, subPphTpp = 0;

        if (res.akumulasi.history && res.akumulasi.history.length > 0) {
            res.akumulasi.history.forEach(h => {
                // Menjumlahkan akumulasi Jan-Nov
                subGaji += h.gaji; subTpp += h.tpp; subIwp += h.iwp; 
                subPphGaji += h.pphGaji; subPphTpp += h.pphTpp;

                tbodyRiwayat.innerHTML += `<tr>
                    <td class="text-center">${h.bulan}</td>
                    <td class="text-end">${fRp(h.gaji)}</td>
                    <td class="text-end">${fRp(h.tpp)}</td>
                    <td class="text-end text-danger">${fRp(h.iwp)}</td>
                    <td class="text-end text-danger">${fRp(h.pphGaji)}</td>
                    <td class="text-end text-danger">${fRp(h.pphTpp)}</td>
                </tr>`;
            });

            // 👇 BARIS SUBTOTAL JANUARI - NOVEMBER 👇
            tbodyRiwayat.innerHTML += `
                <tr class="table-warning border-warning">
                    <td class="text-center fw-bold">SUBTOTAL (JAN-NOV) <br><small class="text-dark"><i>*Dipungut dgn Metode TER</i></small></td>
                    <td class="text-end fw-bold">${fRp(subGaji)}</td>
                    <td class="text-end fw-bold">${fRp(subTpp)}</td>
                    <td class="text-end fw-bold text-danger">${fRp(subIwp)}</td>
                    <td class="text-end fw-bold text-danger">${fRp(subPphGaji)}</td>
                    <td class="text-end fw-bold text-danger">${fRp(subPphTpp)}</td>
                </tr>`;

        } else {
            tbodyRiwayat.innerHTML = `<tr><td colspan="6" class="text-center text-muted py-3">Belum ada riwayat pajak Jan-Nov.</td></tr>`;
        }

        // 👇 BARIS KHUSUS DESEMBER (CLEARING) 👇
        let iwpBulanIni = ((res.gapok || 0) + (res.tjKeluarga || 0)) * 0.0475;
        tbodyRiwayat.innerHTML += `
            <tr class="table-danger border-danger" style="border-width: 2px;">
                <td class="text-center fw-bold text-danger">${globalBulanAktif.toUpperCase()} <br><small class="text-dark"><i>*Dipungut dgn Metode Progresif (Clearing)</i></small></td>
                <td class="text-end fw-bold text-danger">${fRp(res.gajiKotorTER)}</td>
                <td class="text-end fw-bold text-danger">${fRp(res.tppBruto)}</td>
                <td class="text-end fw-bold text-danger">${fRp(iwpBulanIni)}</td>
                <td class="text-end fw-bold text-danger">${fRp(res.pphGajiTER)}</td>
                <td class="text-end fw-bold text-danger">${fRp(res.pph21TKD)}</td>
            </tr>`;
        
        // 👇 BARIS TOTAL KESELURUHAN (SETELAH DESEMBER) 👇
        tbodyRiwayat.innerHTML += `
            <tr class="bg-dark text-white border-dark">
                <td class="text-center fw-bold text-white">TOTAL SETAHUN (Real)</td>
                <td class="text-end fw-bold text-white">${fRp(subGaji + res.gajiKotorTER)}</td>
                <td class="text-end fw-bold text-white">${fRp(subTpp + res.tppBruto)}</td>
                <td class="text-end fw-bold text-white">${fRp(subIwp + iwpBulanIni)}</td>
                <td class="text-end fw-bold text-white">${fRp(subPphGaji + res.pphGajiTER)}</td>
                <td class="text-end fw-bold text-white">${fRp(subPphTpp + res.pph21TKD)}</td>
            </tr>`;
        } else {
            tbodyRiwayat.innerHTML = `<tr><td colspan="6" class="text-center text-muted py-3">Belum ada riwayat pajak.</td></tr>`;
        }

        document.getElementById('desBrutoGaji').innerText = fRp(res.akumulasi.sumBrutoGaji);
        document.getElementById('desBrutoTPP').innerText = fRp(res.akumulasi.sumBrutoTPP);
        document.getElementById('desTotalBruto').innerText = fRp(res.akumulasi.totalBrutoSetahunReal);
        document.getElementById('desBiayaJab').innerText = "- " + fRp(res.akumulasi.biayaJabatanReal);
        document.getElementById('desIWP').innerText = "- " + fRp(res.akumulasi.sumIWP);
        document.getElementById('desNetto').innerText = fRp(res.akumulasi.nettoSetahunReal);
        document.getElementById('desStatusKwn').innerText = res.statusTER;
        document.getElementById('desPTKP').innerText = "- " + fRp(res.ptkp);
        document.getElementById('desPKP').innerText = fRp(res.akumulasi.pkpReal);
        document.getElementById('desPajakSetahun').innerText = fRp(res.pph21Setahun);
        document.getElementById('desPajakDibayar').innerText = "- " + fRp(res.akumulasi.pajakSudahDibayarJanNov);
        document.getElementById('desPphGajiDes').innerText = "- " + fRp(res.pphGajiTER);
        if (document.getElementById('desSisaTerutang')) document.getElementById('desSisaTerutang').innerText = fRp(res.pph21TotalSebulanTER);
        document.getElementById('desPphTKD').innerText = fRp(res.pph21TKD);

    } else {
        
        // BUKA TER, TUTUP DESEMBER
        if(document.getElementById('layoutDesember')) document.getElementById('layoutDesember').classList.add('hidden');
        if(document.getElementById('layoutTER')) document.getElementById('layoutTER').classList.remove('hidden');

        document.getElementById('thTerGajiAwal').innerText = fRp(res.gajiKotorTER);
        document.getElementById('thTerTppBruto').innerText = "+ " + fRp(res.tppBruto);
        document.getElementById('thTerDasar').innerText = fRp(res.dasarPajakTER);
        document.getElementById('thKatTER').innerText = res.katTER;
        document.getElementById('thPctTER').innerText = res.pctTER + "%";
        if(document.getElementById('teksRumusTER')) document.getElementById('teksRumusTER').innerText = `Total PPh21 (${fRp(res.dasarPajakTER)} x ${res.pctTER}%)`;
        document.getElementById('thPphTER').innerText = fRp(res.pph21TotalSebulanTER);
        document.getElementById('thPphGajiLunas').innerText = "- " + fRp(res.pphGajiTER);
        document.getElementById('thPphTKD').innerText = fRp(res.pph21TKD);
                
        // Data Simulasi Progresif
        document.getElementById('thGajiBulanOnly').innerText = fRp(res.gajiKotor); 
        document.getElementById('thGajiTahunOnly').innerText = fRp(res.brutoGajiSetahun); 
        document.getElementById('thBiayaJabGaji').innerText = "- " + fRp(res.biayaJabatanGaji); 
        document.getElementById('thIwpGaji').innerText = "- " + fRp(res.iwpSetahun); 
        document.getElementById('thNettoGaji').innerText = fRp(res.nettoGajiSetahun); 
        document.getElementById('thPtkpGaji').innerText = "- " + fRp(res.ptkp); 
        if(document.getElementById('thStatusKwnGaji')) document.getElementById('thStatusKwnGaji').innerText = res.statusTER; 
        document.getElementById('thPkpGaji').innerText = fRp(res.pkpGaji); 
        document.getElementById('thPphGajiBulan').innerText = fRp(res.pph21GajiSebulanSimulasi); 
        document.getElementById('thGajiBulan').innerText = fRp(res.gajiKotor); 
        document.getElementById('thTppKepgub').innerText = fRp(res.tppKepgub); 
        document.getElementById('thPotPK').innerText = "- " + fRp(res.totalPotPK); 
        document.getElementById('thPotDK').innerText = "- " + fRp(res.totalPotDK); 
        document.getElementById('thTppBrutoBulan').innerText = fRp(res.tppBruto); 
        document.getElementById('thBrutoBulan').innerText = fRp(res.brutoSebulan); 
        document.getElementById('thBruto').innerText = fRp(res.brutoSetahun); 
        document.getElementById('thBiayaJab').innerText = "- " + fRp(res.biayaJabatan); 
        document.getElementById('thIWP').innerText = "- " + fRp(res.iwpSetahun); 
        document.getElementById('thNetto').innerText = fRp(res.nettoSetahun); 
        if(document.getElementById('thStatusKwnTotal')) document.getElementById('thStatusKwnTotal').innerText = res.statusTER; 
        document.getElementById('thPTKP').innerText = "- " + fRp(res.ptkp); 
        document.getElementById('thPKP').innerText = fRp(res.pkp); 
        document.getElementById('thPph5').innerText = fRp(res.pph5); 
        document.getElementById('thPph15').innerText = fRp(res.pph15); 
        document.getElementById('thPph25').innerText = fRp(res.pph25); 
        document.getElementById('thPph30').innerText = fRp(res.pph30); 
        document.getElementById('thPph35').innerText = fRp(res.pph35); 
        document.getElementById('thPajakTahun').innerText = fRp(res.pph21Setahun); 
        document.getElementById('thPajakSebulanProg').innerText = fRp(res.pph21Setahun / 12); 
    }
}

function muatNominatif() {
    if(!objNominatifSetahun) { hitungPajakTahunan(); return; } 
    let res = objNominatifSetahun; 
    document.getElementById('loaderNominatif').classList.add('hidden'); 
    document.getElementById('hasilNominatif').classList.remove('hidden'); 
    
    // 👇 FIX FRONTEND 2: TAMBAH PENJAGA NaN DI MENU 4 👇
    const fRp = (angka) => "Rp " + Math.round(angka || 0).toLocaleString('id-ID');
    
    let isSpesial = (res.jenisPeriode !== "Reguler"); 
    let rasio = res.tppKepgub > 0 ? (res.tppBruto / res.tppKepgub) : 0;
    
    document.getElementById('nomGajiKotor').innerText = fRp(res.gajiKotor); 
    document.getElementById('nomBK').innerText = fRp(res.bk * rasio); 
    document.getElementById('nomPK').innerText = fRp(res.pk * rasio); 
    document.getElementById('nomKK').innerText = fRp(res.kk * rasio); 
    document.getElementById('nomTB').innerText = fRp(res.tb * rasio); 
    document.getElementById('nomKP').innerText = fRp(res.kp * rasio); 
    document.getElementById('nomTppBruto').innerText = fRp(res.tppBruto); 
    document.getElementById('nomTambahanBpjs').innerText = fRp(res.bpjs4);
    
    let tppPlusBpjs = (res.tppBruto || 0) + (res.bpjs4 || 0); 
    document.getElementById('nomTppPlusBpjs').innerText = fRp(tppPlusBpjs);
    
    let infoBebas = isSpesial ? ` <span class="badge bg-danger ms-1">Bebas Pot. ${res.jenisPeriode}</span>` : ""; 
    document.getElementById('nomPotIwp').innerHTML = "- " + fRp(res.iwp1) + infoBebas; 
    document.getElementById('nomPotBpjs').innerHTML = "- " + fRp(res.bpjs4) + infoBebas; 
    document.getElementById('nomPotPph').innerHTML = "- " + fRp(res.pph21TKD); 
    
    let jmlPotongan = (res.iwp1 || 0) + (res.pph21TKD || 0) + (res.bpjs4 || 0); 
    document.getElementById('nomJmlPotongan').innerText = "- " + fRp(jmlPotongan); 
    
    let totalDiterima = tppPlusBpjs - jmlPotongan; 
    document.getElementById('nomTotalDiterima').innerText = fRp(totalDiterima); 
    document.getElementById('nomTppNetto').innerText = fRp(totalDiterima);
    
    let potonganKorpriAktual = globalPotKorpri || 0; 
    document.getElementById('nomTrfBersih').innerText = fRp(totalDiterima); 
    document.getElementById('nomTrfKorpri').innerText = "- " + fRp(potonganKorpriAktual); 
    let masukRekening = totalDiterima - potonganKorpriAktual; 
    document.getElementById('nomMasukRekening').value = fRp(masukRekening);
}

  function navigasiTab(target, force = false) {
    if (!force && globalStatusLock !== "Kunci") { 
        // Hapus peringatan Gaji, karena sudah otomatis disimpen
        if ((target === 'tabTahunan' || target === 'tabNominatif') && !isAbsenTersimpan) { 
            return alertPeringatan("Anda belum menyimpan kehadiran! Silakan klik SIMPAN PERHITUNGAN TPP terlebih dahulu."); 
        } 
    }
    
    if (target === 'tabNominatif' && !objNominatifSetahun) {
        return alertPeringatan("Harap buka menu '3. Penghasilan Setahun / Pajak TER' terlebih dahulu agar sistem dapat mengkalkulasi nilai pajak TPP yang terbaru.");
    }

    // 1. Reset tombol Navigasi (Warna dan Status)
    ['navGaji', 'navAbsen', 'navTahunan', 'navNominatif'].forEach(id => {
        let navBtn = document.getElementById(id);
        if (navBtn) navBtn.classList.remove('active');
    });
    
    // 2. Aktifkan HANYA tombol Navigasi yang diklik
    let navId = "nav" + target.replace("tab", ""); 
    let targetNavBtn = document.getElementById(navId);
    if (targetNavBtn) targetNavBtn.classList.add('active'); 

    // 3. PAKSA SEMBUNYI semua isi tab menggunakan class .hidden buatan Anda
    ['tabGaji', 'tabAbsen', 'tabTahunan', 'tabNominatif'].forEach(id => {
        let tab = document.getElementById(id);
        if (tab) {
            tab.classList.remove('show', 'active');
            tab.classList.add('hidden'); // Eksekusi lenyap tanpa sisa ruang
        }
    });

    // 4. Tampilkan HANYA isi tab yang dituju
    let targetTabPane = document.getElementById(target);
    if (targetTabPane) {
        targetTabPane.classList.remove('hidden');
        
        // Jeda sangat singkat (10ms) agar DOM browser selesai mereset tinggi (height) layout
        // sebelum animasi dari Bootstrap dipanggil ulang.
        setTimeout(() => {
            targetTabPane.classList.add('show', 'active');
        }, 10);
    }
    
    // 5. Muat fungsi pendukung untuk hitungan otomatis
    // [PERBAIKAN] Tambahkan && !objNominatifSetahun agar tidak loading 2x
    if (target === 'tabTahunan' && !objNominatifSetahun) hitungPajakTahunan(); 
    if (target === 'tabNominatif') muatNominatif();
}

  function bukaModalImport(jenis) { document.getElementById('importJenis').value = jenis; document.getElementById('fileImport').value = ""; document.getElementById('importModalTitle').innerText = jenis === 'pegawai' ? "Import Data Pegawai (Excel)" : "Import Data (Excel)"; 
    let modalObj = bootstrap.Modal.getInstance(document.getElementById('modalImportExcel')) || new bootstrap.Modal(document.getElementById('modalImportExcel'));
    modalObj.show(); 
  }

  async function unduhTemplateExcel() {
      try {
          let jenis = document.getElementById('importJenis').value; 
          const workbook = new ExcelJS.Workbook();
          let worksheet, wsData, sheetName, fileName;

          // 1. Tentukan Header & Nama File Berdasarkan Jenis
          if (jenis === 'pegawai') { 
              wsData = ["NIP (Angka Saja)", "Nama Lengkap", "Tanggal Lahir (YYYY-MM-DD)", "Golongan (Contoh: III/b)", "Unit Kerja", "Nama Jabatan (Sesuai Master)", "Jenis Jabatan (Struktural / Fungsional Tertentu / Fungsional Umum)", "Status Kawin (Contoh: K/1 = 3)", "Gaji Pokok", "Tunjangan Jabatan", "No. Rekening"]; 
              sheetName = "Template_Pegawai"; 
              fileName = "Template_Import_Pegawai.xlsx"; 
          } 
          else if(jenis === 'pergub') { 
        // 👇 TAMBAH KOLOM STATUS PEGAWAI DI SINI 👇
        wsData = ["Nama Jabatan", "Kelas Jabatan", "Beban Kerja (BK)", "Prestasi Kerja (PK)", "Kondisi Kerja (KK)", "Tempat Bertugas (TB)", "Kelangkaan Profesi (KP)", "Status Pegawai (PNS / PPPK)"]; 
        sheetName = "Template_Master_TPP"; 
        fileName = "Template_Import_Master_TPP.xlsx"; 
          } 
          else if (jenis === 'akun') { 
              wsData = ["Username", "Password", "Role (Admin / Operator)", "Unit Kerja (Harus sesuai ketikan DB_Unit_Kerja)", "Email"]; 
              sheetName = "Template_Akun"; 
              fileName = "Template_Import_Akun.xlsx"; 
          }

          // 2. Buat Sheet dan Masukkan Header
          worksheet = workbook.addWorksheet(sheetName);
          worksheet.addRow(wsData);

          // 3. Desain Header Biar Cantik (Warna Biru, Teks Putih Tebal)
          worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
          worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2980B9' } };
          worksheet.columns.forEach(column => { column.width = 25; }); // Lebarkan kolom

          // 4. Proses Download (Pakai ExcelJS & FileSaver)
          const buffer = await workbook.xlsx.writeBuffer();
          const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
          saveAs(blob, fileName); // saveAs ini dari library FileSaver.js di HTML Anda
          
      } catch (error) {
          alertError("Gagal membuat template: " + error.message);
      }
  }

  // =========================================================
  // FUNGSI: IMPORT EXCEL (BKD / MASTER / AKUN) - AUTO REFRESH
  // =========================================================
  async function prosesImportExcel(fileInputElemen = null, periodeTarget = null, periodeDataTerbaru = null, btnSubmitModal = null) {
      let fileInput = fileInputElemen || document.getElementById('fileImport'); 
      if(!fileInput.files[0]) return alertPeringatan("Pilih file Excel terlebih dahulu!");
      
      let jenis = document.getElementById('importJenis').value; 
      let file = fileInput.files[0]; 
      
      // Jika dipanggil dari Modal Tambah Periode, paksa mode "pegawai"
      if (periodeTarget) jenis = 'pegawai';
      let bulanTarget = periodeTarget || globalBulanAktif;

      startLoading("Membaca File Excel & Mengirim ke Database...");
      
      try {
          let arrayBuffer = await file.arrayBuffer();
          const wb = new ExcelJS.Workbook();
          await wb.xlsx.load(arrayBuffer);
          const worksheet = wb.worksheets[0];
          
          let data2D = [];
          worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
              let r = [];
              for(let i = 1; i <= Math.max(row.cellCount, 40); i++) {
                  let cell = row.getCell(i);
                  let val = cell.value;
                  if(val && typeof val === 'object') {
                      if(val.result !== undefined) val = val.result;
                      else if(val.richText) val = val.richText.map(t => t.text).join("");
                  }
                  r.push(val === null || val === undefined ? "" : val);
              }
              data2D.push(r);
          });

          if(data2D.length === 0) {
              stopLoading(); 
              if(btnSubmitModal) btnSubmitModal.disabled = false;
              return alertError("File Excel kosong atau format tidak sesuai!");
          }

          let payload = [];

          if(jenis === 'pegawai') {
              if(!bulanTarget) { stopLoading(); return alertError("Pilih Periode Bulan terlebih dahulu!"); }

              for(let i = 1; i < data2D.length; i++) {
                  let row = data2D[i];
                  let nipVal = row[3]; 
                  if(!nipVal) continue; 

                  let nipBersih = String(nipVal).replace(/[\s-']/g, '');
                  if(!/^\d{18}$/.test(nipBersih)) continue; 

                  let jenisPegawaiRaw = String(row[31] || "").trim().toUpperCase();
                  let statusPegawaiVal = "PNS"; 
                  if (jenisPegawaiRaw.includes("PPPK") || jenisPegawaiRaw.includes("P3K")) { statusPegawaiVal = "PPPK"; }

                  payload.push({
                  nip: nipBersih, 
                  nama: String(row[4] || "").trim(),              
                  unitkerja: String(row[8] || "").trim(),       
                  unorInduk: String(row[10] || "").trim(),    
                  skp: String(row[16] || "").trim(),          
                  golongan: String(row[30] || "").trim(),       
                  statusPegawai: statusPegawaiVal
                  // Atribut lain sengaja DIHAPUS agar backend tahu bahwa kita TIDAK ingin menimpanya!
                });
              }

              if(payload.length === 0) {
                  stopLoading(); 
                  if(btnSubmitModal) btnSubmitModal.disabled = false;
                  return alertPeringatan("Tidak ada data NIP valid (18 Digit) di Kolom D (Excel).");
              }

              Swal.getHtmlContainer().innerHTML = `Menyimpan ${payload.length} Pegawai ke Server...`;
              let resImport = await fetchAPI("importPegawaiMassal", {payload: payload, bulanAktif: bulanTarget});
              
              stopLoading(); 
              if(btnSubmitModal) btnSubmitModal.disabled = false;
              
              if(String(resImport).includes("Error")) { 
                  alertError(resImport.pesan || resImport); 
              } else { 
                  alertSukses(`Data SKP berhasil diimpor ke bulan ${bulanTarget}!`); 
              }

              // JIKA BERASAL DARI MODAL TAMBAH PERIODE (AUTO REFRESH!)
              if (periodeTarget && periodeDataTerbaru) {
                  let modalObj = bootstrap.Modal.getInstance(document.getElementById('modalTambahPeriode'));
                  if(modalObj) modalObj.hide();
                  
                  renderDropdownPeriode(periodeDataTerbaru);
                  document.getElementById('pilihPeriodeUtama').value = bulanTarget;
                  document.getElementById('formTambahPeriode').reset();
                  toggleSumberData();
                  
                  window.cacheDataPegawaiAll = null; 
                  globalBulanAktif = bulanTarget;
                  sessionStorage.setItem('globalBulanAktif', globalBulanAktif);
                  masukAplikasi(); // Langsung pindah ke tabel pegawai
              } 
              // JIKA BERASAL DARI MENU IMPORT BIASA
              else {
                  let modalObj = bootstrap.Modal.getInstance(document.getElementById('modalImportExcel'));
                  if(modalObj) modalObj.hide();
                  muatDataPegawai(true); // Paksa narik data baru dari server
              }
          } 
          else if(jenis === 'pergub') {
              for(let i = 1; i < data2D.length; i++) { 
                  let row = data2D[i]; 
                  if(!row[0]) continue; 
                  
                  payload.push({ 
                      namaJabatan: row[0], 
                      kelasJabatan: row[1], 
                      bk: parseFloat(row[2])||0, 
                      pk: parseFloat(row[3])||0, 
                      kk: parseFloat(row[4])||0, 
                      tb: parseFloat(row[5])||0, 
                      kp: parseFloat(row[6])||0,
                      statusPegawai: String(row[7] || "PNS").trim() // 👇 TANGKAP KOLOM KE 8 (Indeks 7)
                  }); 
              }
              let res = await fetchAPI("importPergubMassal", payload); 
              stopLoading(); 
              
              // 👇 Perbaiki bacaan respon errornya
              if(res && res.status === "error") { 
                  alertError(res.pesan); 
              } else { 
                  alertSukses(res.pesan || "Sukses"); 
                  bootstrap.Modal.getInstance(document.getElementById('modalImportExcel')).hide(); 
                  muatDataPergub(); 
              }
          } 
          else if(jenis === 'akun') {
              for(let i = 1; i < data2D.length; i++) { 
                  let row = data2D[i]; 
                  if(!row[0]) continue; 
                  payload.push({ username: row[0], password: row[1], role: row[2], unitkerja: row[3], email: row[4] }); 
              }
              let res = await fetchAPI("importAkunMassal", payload); stopLoading(); 
              if(String(res).includes("Error") || String(res).includes("Gagal")) alertError(res.pesan || res); else alertSukses(res); bootstrap.Modal.getInstance(document.getElementById('modalImportExcel')).hide(); muatDaftarAkun(); 
          }
      } catch(e) {
          stopLoading();
          if(btnSubmitModal) btnSubmitModal.disabled = false;
          alertError("Gagal memproses file Excel: " + e.message);
      }
  }

  let globalDataAkun = []; let currentAkunPage = 1; let akunRowsPerPage = 15;

  async function muatDaftarAkun(forceRefresh = false) {
    switchView('viewManajemenAkun'); 
    
    if (!forceRefresh && globalDataAkun && globalDataAkun.length > 0) {
      terapkanFilterAkun();
      return;
    }
    
    startLoading("Memuat Data Akun...");
    let data = await fetchAPI("getDaftarAkun", {}); 
    stopLoading(); 
    
    if(data && !data.error) { 
      globalDataAkun = data; 
      try {
          let currentCache = JSON.parse(localStorage.getItem('simTppCacheData') || "{}");
          currentCache.akun = data;
          localStorage.setItem('simTppCacheData', JSON.stringify(currentCache));
      } catch(e) {}
      terapkanFilterAkun(); 
    } else { 
      alertError(data.pesan || "Gagal memuat daftar akun dari server!"); 
    }
  }

  function terapkanFilterAkun() { 
      let cari = document.getElementById('filterCariAkun').value.toLowerCase(); 
      let filtered = globalDataAkun.filter(row => { 
          let roleAkun = String(row[2]).trim();
          
          // 👇 KUNCI RAHASIA: Kalau yang login Admin OPD, Super Admin otomatis Ghaib!
          if (currentUser.role === "Admin OPD" && roleAkun === "Super Admin") {
              return false; 
          }

          return String(row[0]).toLowerCase().includes(cari) || String(row[3]).toLowerCase().includes(cari); 
      }); 
      renderTabelAkun(filtered); 
  }

  function renderTabelAkun(data) {
    let tbody = document.getElementById('tabelBodyAkun'); if(!data.length) { document.getElementById('infoPaginationAkun').innerText = "Menampilkan 0 data"; document.getElementById('btnPaginationAkun').innerHTML = ""; return tbody.innerHTML = '<tr><td colspan="5" class="text-center text-danger fw-bold">Data tidak ditemukan.</td></tr>'; }
    let totalPages = Math.ceil(data.length / akunRowsPerPage); if(currentAkunPage > totalPages) currentAkunPage = totalPages; let start = (currentAkunPage - 1) * akunRowsPerPage; let end = start + akunRowsPerPage; let paginatedData = data.slice(start, end); let html = "";
    paginatedData.forEach(r => { let u = escapeStr(r[0]); let role = escapeStr(r[2]); let unit = escapeStr(r[3]); let email = escapeStr(r[4]); let uuid = escapeStr(r[5]); let badge = role === 'Admin' ? 'bg-danger' : 'bg-primary'; let args = `'edit', '${u}', '${role}', '${unit}', '${email}', '${uuid}'`; html += `<tr><td><b>${u}</b></td><td>${unit}</td><td><span class="badge ${badge}">${role}</span></td><td>${email}</td><td class="text-center"><button class="btn btn-sm btn-info text-white me-1 fw-bold" onclick="resetPasswordAkun('${uuid}', '${u}')"><i class="bi bi-key"></i></button><button class="btn btn-sm btn-warning me-1" onclick="bukaModalAkun(${args})"><i class="bi bi-pencil"></i></button><button class="btn btn-sm btn-danger" onclick="hapusAkun('${uuid}', '${u}')"><i class="bi bi-trash"></i></button></td></tr>`; });
    tbody.innerHTML = html; document.getElementById('infoPaginationAkun').innerText = `Tampil ${start + 1} - ${Math.min(end, data.length)} dari ${data.length}`;
    let btnHtml = `<button class="btn btn-sm btn-outline-primary" onclick="ubahHalamanAkun(${currentAkunPage - 1})" ${currentAkunPage === 1 ? 'disabled' : ''}>Mundur</button><span class="btn btn-sm btn-primary disabled text-white fw-bold">Hal ${currentAkunPage}/${totalPages}</span><button class="btn btn-sm btn-outline-primary" onclick="ubahHalamanAkun(${currentAkunPage + 1})" ${currentAkunPage === totalPages ? 'disabled' : ''}>Maju</button>`; document.getElementById('btnPaginationAkun').innerHTML = btnHtml;
  }
  function ubahHalamanAkun(page) { currentAkunPage = page; terapkanFilterAkun(); }

  let arrayUnitKerjaFull = [];
  let arraySubUnitValid = [];

  function terapkanDataInit(dataInit, cUser, bAktif, jASN, isSilentUpdate = false) {
    globalSettingsCache = dataInit.setting;
    let res = dataInit.setting;
    if(res) {
      document.getElementById('loginNamaInstansi').innerText = res.Nama_Instansi || "PEMERINTAH PROVINSI JAMBI"; 
      document.getElementById('loginNamaDinas').innerText = res.Nama_Dinas || "DINAS PENDIDIKAN"; 
      document.getElementById('loginSubDinas').innerText = res.Nama_Sub_Dinas ? res.Nama_Sub_Dinas.toUpperCase() : "";
      if(document.getElementById('mOpdInduk')) document.getElementById('mOpdInduk').value = res.Nama_Dinas || "";
      if(res.Logo_Instansi) { let pI = document.getElementById('loginLogoInstansi'); pI.src = res.Logo_Instansi; pI.classList.remove('hidden'); let mLogo = document.getElementById('mHeaderLogo'); if(mLogo) { mLogo.src = res.Logo_Instansi; mLogo.classList.remove('hidden'); } }
      if(res.Logo_Dinas) { let pD = document.getElementById('loginLogoDinas'); pD.src = res.Logo_Dinas; pD.classList.remove('hidden'); }
      globalPotKorpri = parseFloat(res.Pot_Korpri) || 0; 
    }

    renderDropdownPeriode(dataInit.periode);
    
    arrayUnitKerjaFull = dataInit.unitKerjaFull || []; 
    arrayUnitKerjaValid = arrayUnitKerjaFull.map(u => u.nama); 
    arraySubUnitValid = [...new Set(arrayUnitKerjaFull.map(u => u.subUnit).filter(Boolean))]; 

    setupAutocomplete(document.getElementById('mUnitKerja'), arrayUnitKerjaValid);
    setupAutocomplete(document.getElementById('uUnitKerjaOp'), arrayUnitKerjaValid);
    
    globalDataPergub = dataInit.pergub || [];
    if(!isSilentUpdate) terapkanFilterPergub(); 

    if(cUser) {
      currentUser = cUser; 
      document.getElementById('navInfoRole').innerText = currentUser.unitkerja;
      
      if (dataInit.akun && dataInit.akun.length > 0) {
          globalDataAkun = dataInit.akun;
      }
      
      let r = currentUser.role;
      let isSuper = (r === "Super Admin");
      let isOpd = (r === "Admin OPD");

      // 1. MASTER TPP: Hanya Super Admin
      if(isSuper) { 
          document.getElementById('btnMasterAdmin').classList.remove('hidden'); 
      } else { 
          document.getElementById('btnMasterAdmin').classList.add('hidden'); 
      }
      
      // 2. KELOLA AKUN: Super Admin DAN Admin OPD bisa lihat!
      if(isSuper || isOpd) { 
          document.getElementById('btnAkunAdmin').classList.remove('hidden'); 
      } else { 
          document.getElementById('btnAkunAdmin').classList.add('hidden'); 
      }
      
      // 3. SETTING: Hanya Super Admin & Admin OPD (Operator disembunyikan)
      if(isSuper || isOpd) { 
          document.getElementById('btnSettingAdmin').classList.remove('hidden'); 
          if(document.getElementById('mHeaderAdminIcons')) document.getElementById('mHeaderAdminIcons').classList.remove('hidden');
      } else { 
          document.getElementById('btnSettingAdmin').classList.add('hidden'); 
          if(document.getElementById('mHeaderAdminIcons')) document.getElementById('mHeaderAdminIcons').classList.add('hidden');
      }
      
      // 4. IKON DI HP MOBILE
      let iconMasterHP = document.querySelector('i[onclick="switchView(\'viewMasterPergub\')"]');
      let iconAkunHP = document.querySelector('i[onclick="muatDaftarAkun()"]');

      if(isSuper || isOpd) { 
          document.getElementById('btnTambahBulan').classList.remove('hidden'); 
          document.getElementById('btnKelolaBulan').classList.remove('hidden'); 
          
          // Ikon Master TPP HP: Hanya Super Admin
          if(iconMasterHP && isSuper) iconMasterHP.classList.remove('hidden'); 
          else if(iconMasterHP) iconMasterHP.classList.add('hidden');
          
          // Ikon Kelola Akun HP: Super Admin & Admin OPD
          if(iconAkunHP && (isSuper || isOpd)) iconAkunHP.classList.remove('hidden'); 
          else if(iconAkunHP) iconAkunHP.classList.add('hidden');
      } else { 
          document.getElementById('btnTambahBulan').classList.add('hidden'); 
          document.getElementById('btnKelolaBulan').classList.add('hidden'); 
          if(iconMasterHP) iconMasterHP.classList.add('hidden');
          if(iconAkunHP) iconAkunHP.classList.add('hidden');
      }
      document.getElementById('viewLogin').classList.add('hidden'); 
      document.getElementById('viewLanding').classList.add('hidden'); 
      document.getElementById('mainNav').classList.remove('hidden');
      
      globalBulanAktif = bAktif; 
      globalHariKerja = parseInt(sessionStorage.getItem('globalHariKerja')) || 22; 
      globalJenisASN = jASN; 
      
      if(globalBulanAktif) { 
          document.getElementById('labelBulanAktif').innerText = globalBulanAktif + (globalJenisASN ? " (" + globalJenisASN + ")" : ""); 
          document.getElementById('labelHKAktif').innerText = globalHariKerja; 
          if(!isSilentUpdate) updateFilterGolDropdown(); 
      }
      
      if(!isSilentUpdate) {
          switchView('viewPilihBulan'); 
          document.getElementById('btnMenuPegawai').disabled = true; 
          toggleTombolPegawaiHP(false); 
      }
      
      if(globalBulanAktif && dataInit.pegawai) { 
          window.cacheDataPegawaiAll = dataInit.pegawai; 
          window.cacheDataPegawaiBulan = globalBulanAktif; 
          
          globalDataPegawai = window.cacheDataPegawaiAll.filter(p => String(p[14] || "PNS").toUpperCase() === globalJenisASN);
          
          let unitList = [...new Set(globalDataPegawai.map(item => item[5]).filter(Boolean))]; 
          let unitDropdown = document.getElementById('filterUnitKerja'); 
          
          let filterValue = unitDropdown.value; 
          unitDropdown.innerHTML = '<option value="">Semua Unit Kerja</option>'; 
          unitList.forEach(s => unitDropdown.innerHTML += `<option value="${s}">${s}</option>`); 
          unitDropdown.value = filterValue;

          terapkanFilter();
      }
    } else {
      if(!isSilentUpdate) {
          document.getElementById('viewLanding').classList.add('hidden'); 
          document.getElementById('viewLogin').classList.remove('hidden'); 
          document.body.classList.add('bg-gradient-login');
      }
    }
  }

  function toggleUnitKerjaAkun() { 
    let r = document.getElementById('uRole').value; 
    let op = document.getElementById('uUnitKerjaOp'); 
    let adm = document.getElementById('uUnitKerjaAdmin'); 
    let btnSimpan = document.getElementById('btnSimpanAkun');
    
    if(r === 'Super Admin' || r === 'Admin OPD') { 
        if(op && op.parentNode) op.parentNode.classList.add('hidden'); 
        if(adm) { adm.classList.remove('hidden'); adm.required = true; }
        
        let namaDinas = document.getElementById('loginNamaDinas').innerText;
        let namaInstansi = document.getElementById('loginNamaInstansi').innerText || "PEMERINTAH PROVINSI JAMBI";
        
        if(adm) adm.value = (r === 'Super Admin') ? namaInstansi : namaDinas;
        if(btnSimpan) btnSimpan.disabled = false;
        
    } else if (r === 'Admin Sub Unit/UPTD') {
        if(adm) { adm.classList.add('hidden'); adm.required = false; }
        if(op && op.parentNode) op.parentNode.classList.remove('hidden'); 
        
        setupAutocomplete(document.getElementById('uUnitKerjaOp'), arraySubUnitValid);
        
        let newOp = document.getElementById('uUnitKerjaOp'); 
        if(newOp) newOp.placeholder = "Ketik nama Sub Unit/UPTD/Kab/Kota...";
        
        if(arraySubUnitValid.length === 0) {
            if(btnSimpan) btnSimpan.disabled = true;
        } else { 
            if(btnSimpan) btnSimpan.disabled = false; 
        }
        
    } else { 
        if(adm) { adm.classList.add('hidden'); adm.required = false; }
        if(op && op.parentNode) op.parentNode.classList.remove('hidden'); 
        
        setupAutocomplete(document.getElementById('uUnitKerjaOp'), arrayUnitKerjaValid);
        
        let newOp = document.getElementById('uUnitKerjaOp'); 
        if(newOp) {
           newOp.placeholder = "Ketik untuk mencari unit kerja sekolah/UPTD...";
           if(arrayUnitKerjaValid.length === 0) newOp.placeholder = "DATA UNIT KERJA KOSONG!";
        }
        
        if(arrayUnitKerjaValid.length === 0) {
            if(btnSimpan) btnSimpan.disabled = true;
        } else { 
            if(btnSimpan) btnSimpan.disabled = false; 
        }
    } 
  }

  function bukaModalAkun(aksi, user="", role="Operator", unit="", email="", uuid="") {
    document.getElementById('uAksi').value = aksi; 
    document.getElementById('uUser').value = user; 
    document.getElementById('uEmail').value = email; 
    document.getElementById('uUuid').value = uuid;
    
    let optRole = document.getElementById('uRole');
    
    // 👇 Sembunyikan Opsi Super Admin dari Dropdown jika yang login Admin OPD
    for (let i = 0; i < optRole.options.length; i++) {
        if (optRole.options[i].value === "Super Admin") {
            optRole.options[i].style.display = (currentUser.role === "Admin OPD") ? "none" : "block";
        }
    }
    
    // Setel default role agar tidak error
    if (currentUser.role === "Admin OPD" && role === "Super Admin") role = "Operator";
    optRole.value = role; 
    
    let passInp = document.getElementById('uPass'); 
    if(aksi === 'edit') { 
        passInp.required = false; passInp.value = ""; 
        document.getElementById('uPassHint').innerText = "(Kosongkan jika tak ingin ganti)"; 
        document.getElementById('akunModalTitle').innerText = "Edit Akun Pengguna"; 
    } else { 
        passInp.required = true; passInp.value = ""; 
        document.getElementById('uPassHint').innerText = ""; 
        document.getElementById('akunModalTitle').innerText = "Tambah Akun Baru"; 
    } 
    
    toggleUnitKerjaAkun(); 

    let op = document.getElementById('uUnitKerjaOp');
    if(role === 'Super Admin' || role === 'Admin OPD') { 
        if(op) op.value = ""; 
    } else { 
        if(op) op.value = unit; 
    }
    
    let modalObj = bootstrap.Modal.getInstance(document.getElementById('modalKelolaAkun')) || new bootstrap.Modal(document.getElementById('modalKelolaAkun'));
    modalObj.show();
  }

  async function submitAkun(e) {
    e.preventDefault(); 
    let aksi = document.getElementById('uAksi').value; 
    let r = document.getElementById('uRole').value; 
    
    let s = (r === 'Super Admin' || r === 'Admin OPD') ? document.getElementById('uUnitKerjaAdmin').value : document.getElementById('uUnitKerjaOp').value.trim(); 
    
    if(r === 'Operator' && !arrayUnitKerjaValid.includes(s)) { 
        return alertPeringatan("Unit Kerja tidak valid! Silakan pilih dari daftar dropdown yang muncul saat Anda mengetik."); 
    }
    
    startLoading("Menyimpan Akun..."); 
    let d = { username: document.getElementById('uUser').value, password: document.getElementById('uPass').value, role: r, unitkerja: s, email: document.getElementById('uEmail').value, uuid: document.getElementById('uUuid').value };
    
    let res = await fetchAPI(aksi === 'edit' ? "updateAkun" : "simpanAkun", d); 
    stopLoading(); 
    
    if(String(res).includes("Error")) { 
        alertError(res.pesan || res); 
    } else { 
        alertSukses(res); 
        bootstrap.Modal.getInstance(document.getElementById('modalKelolaAkun')).hide(); 
        muatDaftarAkun(true); 
    }
  }

  function hapusAkun(uuid, user) { 
    Swal.fire({
      title: 'Hapus Akun Pengguna?',
      text: `Yakin ingin menghapus akun: ${user}? (Data tidak bisa dikembalikan!)`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#dc3545',
      cancelButtonColor: '#6c757d',
      confirmButtonText: '<i class="bi bi-trash"></i> Ya, Hapus!'
    }).then(async (result) => {
      if (result.isConfirmed) {
        startLoading("Menghapus..."); 
        let res = await fetchAPI("hapusAkun", uuid); 
        stopLoading(); 
        if(String(res).includes("Error")) { alertError(res.pesan || res); } 
        else { alertSukses(res); muatDaftarAkun(true); } 
      }
    });
  }

  function resetPasswordAkun(uuid, user) { 
    Swal.fire({
      title: 'Reset Password?',
      text: `Yakin ingin mereset password akun "${user}" menjadi standar "123456" ?`,
      icon: 'question',
      showCancelButton: true,
      confirmButtonColor: '#ffc107',
      cancelButtonColor: '#6c757d',
      confirmButtonText: '<i class="bi bi-key"></i> Ya, Reset!'
    }).then(async (result) => {
      if (result.isConfirmed) {
        startLoading("Mereset Password..."); 
        let res = await fetchAPI("resetPasswordAdmin", uuid); 
        stopLoading(); 
        if(String(res).includes("SUKSES")) {
            alertSukses(res); 
            muatDaftarAkun(true); 
        } else {
            alertError(res.pesan || res);
        }
      }
    });
  }

  function bukaModalProfil() { 
      document.getElementById('myUser').value = currentUser.username || ""; 
      document.getElementById('myEmail').value = currentUser.email || ""; 
      document.getElementById('myPass').value = ""; 
      let modalObj = bootstrap.Modal.getInstance(document.getElementById('modalProfil')) || new bootstrap.Modal(document.getElementById('modalProfil'));
      modalObj.show(); 
  }

  async function submitProfilSendiri(e) { e.preventDefault(); startLoading("Update Profil..."); let d = { uuid: currentUser.uuid, username: currentUser.username, role: currentUser.role, unitkerja: currentUser.unitkerja, email: document.getElementById('myEmail').value, password: document.getElementById('myPass').value }; let res = await fetchAPI("updateAkun", d); stopLoading(); if(String(res).includes("Error")) { alertError(res.pesan || res); } else { alertSukses("Profil berhasil diupdate! Jika ganti password, ingat untuk login selanjutnya."); bootstrap.Modal.getInstance(document.getElementById('modalProfil')).hide(); } }

  function bukaLupaPassword() { Swal.fire({ title: 'Lupa Password?', text: "Masukkan alamat email yang terdaftar pada akun Anda.", input: 'email', showCancelButton: true, confirmButtonText: 'Kirim', cancelButtonText: 'Batal', showLoaderOnConfirm: true, preConfirm: async (email) => { let res = await fetchAPI("prosesLupaPassword", email); return res; } }).then((result) => { if (result.isConfirmed) { if(String(result.value).includes("SUKSES")) alertSukses(result.value); else alertError(result.value.pesan || result.value); } }); }

  let globalSettingsCache = null; 

  function isiFormSetting(res) {
    document.getElementById('sDinas').value = res.Nama_Dinas || ""; 
    document.getElementById('sSubDinas').value = res.Nama_Sub_Dinas || ""; 
    document.getElementById('sInstansi').value = res.Nama_Instansi || ""; 
    document.getElementById('sKepJab').value = res.Kepala_Jabatan || ""; 
    document.getElementById('sKepNama').value = res.Kepala_Nama || ""; 
    document.getElementById('sKepPangkat').value = res.Kepala_Pangkat || ""; 
    document.getElementById('sKepNIP').value = res.Kepala_NIP || ""; 
    document.getElementById('sBenNama').value = res.Bendahara_Nama || ""; 
    document.getElementById('sBenPangkat').value = res.Bendahara_Pangkat || ""; 
    document.getElementById('sBenNIP').value = res.Bendahara_NIP || ""; 
    document.getElementById('sRekKorpri').value = res.Rek_Korpri || ""; 
    document.getElementById('sPotKorpri').value = res.Pot_Korpri || ""; 
    document.getElementById('sLogoDinas').value = res.Logo_Dinas || ""; 
    document.getElementById('sLogoInstansi').value = res.Logo_Instansi || "";
    
    let pD = document.getElementById('previewLogoDinas');
    if(res.Logo_Dinas) { pD.src = res.Logo_Dinas; pD.classList.remove('hidden'); } else { pD.classList.add('hidden'); }
    
    let pI = document.getElementById('previewLogoInstansi');
    if(res.Logo_Instansi) { pI.src = res.Logo_Instansi; pI.classList.remove('hidden'); } else { pI.classList.add('hidden'); }

    let role = currentUser.role;
    let boxSub = document.getElementById('sSubDinas').parentNode; 
    let labelSub = boxSub.querySelector('label'); 
    
    let inpInstansi = document.getElementById('sInstansi');
    let inpDinas = document.getElementById('sDinas');
    let inpSubDinas = document.getElementById('sSubDinas');
    let inpRek = document.getElementById('sRekKorpri');
    let inpPot = document.getElementById('sPotKorpri');
    
    let btnLogoInstansi = document.getElementById('fileLogoInstansi');
    let btnHapusLogoInstansi = btnLogoInstansi.nextElementSibling; 
    let btnLogoDinas = document.getElementById('fileLogoDinas');
    let btnHapusLogoDinas = btnLogoDinas.nextElementSibling; 
    
    [inpInstansi, inpDinas, inpRek, inpPot, inpSubDinas].forEach(el => {
        el.readOnly = false;
        el.classList.remove('readonly-field', 'bg-light');
    });
    [btnLogoInstansi, btnHapusLogoInstansi, btnLogoDinas, btnHapusLogoDinas].forEach(el => el.classList.remove('hidden'));
    
    if (role === "Super Admin") {
        boxSub.classList.add('hidden'); 
    } 
    else if (role === "Admin OPD") {
        boxSub.classList.add('hidden'); 
        inpInstansi.readOnly = true; inpInstansi.classList.add('readonly-field', 'bg-light');
        btnLogoInstansi.classList.add('hidden'); btnHapusLogoInstansi.classList.add('hidden');
    } 
    else if (role === "Operator" || role === "Admin Sub Unit/UPTD") {
        boxSub.classList.remove('hidden'); 
        labelSub.innerText = role === "Operator" ? "Unit Kerja" : "Sub Unit / UPTD"; 
        inpSubDinas.value = currentUser.unitkerja; 
        
        [inpInstansi, inpDinas, inpSubDinas, inpRek, inpPot].forEach(el => {
            el.readOnly = true; el.classList.add('readonly-field', 'bg-light');
        });
        [btnLogoInstansi, btnHapusLogoInstansi, btnLogoDinas, btnHapusLogoDinas].forEach(el => el.classList.add('hidden'));
    }
  }

  async function submitSetting(e) {
    e.preventDefault(); startLoading("Menyimpan Pengaturan..."); 
    
    let subDinasVal = document.getElementById('sSubDinas').value || "";
    if (currentUser.role !== "Operator") subDinasVal = ""; 
    
    let d = { usernameAksi: currentUser.username, Nama_Dinas: document.getElementById('sDinas').value, Nama_Sub_Dinas: subDinasVal, Nama_Instansi: document.getElementById('sInstansi').value, Kepala_Jabatan: document.getElementById('sKepJab').value, Kepala_Nama: document.getElementById('sKepNama').value, Kepala_Pangkat: document.getElementById('sKepPangkat').value, Kepala_NIP: document.getElementById('sKepNIP').value, Bendahara_Nama: document.getElementById('sBenNama').value, Bendahara_Pangkat: document.getElementById('sBenPangkat').value, Bendahara_NIP: document.getElementById('sBenNIP').value, Rek_Korpri: document.getElementById('sRekKorpri').value, Pot_Korpri: document.getElementById('sPotKorpri').value || 0, Logo_Dinas: document.getElementById('sLogoDinas').value || "", Logo_Instansi: document.getElementById('sLogoInstansi').value || "" };
    
    let res = await fetchAPI("simpanSetting", d); 
    stopLoading(); 
    
    if(String(res).includes("Akses Ditolak") || String(res).includes("Error")) { 
        alertError(res.pesan || res); 
    } else { 
        alertSukses(res); 
        globalSettingsCache = d; 
        
        try {
            let currentCache = JSON.parse(localStorage.getItem('simTppCacheData') || "{}");
            currentCache.setting = d;
            localStorage.setItem('simTppCacheData', JSON.stringify(currentCache));
        } catch(e) {}
        
        document.getElementById('loginNamaInstansi').innerText = d.Nama_Instansi || "PEMERINTAH PROVINSI JAMBI";
        document.getElementById('loginNamaDinas').innerText = d.Nama_Dinas || "DINAS PENDIDIKAN";
        document.getElementById('loginSubDinas').innerText = d.Nama_Sub_Dinas ? d.Nama_Sub_Dinas.toUpperCase() : "";
        globalPotKorpri = parseFloat(d.Pot_Korpri) || 0;
    }
  }

  async function bukaSetting() {
    switchView('viewSetting');
    
    if (globalSettingsCache) {
        isiFormSetting(globalSettingsCache);
    } else {
        startLoading("Memuat Pengaturan...");
        let res = await fetchAPI("getSetting", {});
        stopLoading();
        if(res && !res.error) {
            globalSettingsCache = res;
            isiFormSetting(res);
        }
    }
  }

  function hapusLogo(tipe) {
      if (tipe === 'Instansi') { document.getElementById('sLogoInstansi').value = ""; document.getElementById('previewLogoInstansi').src = ""; document.getElementById('previewLogoInstansi').classList.add('hidden'); document.getElementById('fileLogoInstansi').value = ""; } 
      else { document.getElementById('sLogoDinas').value = ""; document.getElementById('previewLogoDinas').src = ""; document.getElementById('previewLogoDinas').classList.add('hidden'); document.getElementById('fileLogoDinas').value = ""; }
  }

  function handleLogoUpload(input, hiddenId, previewId) { const file = input.files[0]; if(!file) return; const reader = new FileReader(); reader.readAsDataURL(file); reader.onload = (e) => { const img = new Image(); img.src = e.target.result; img.onload = () => { const canvas = document.createElement('canvas'); const ctx = canvas.getContext('2d'); const size = Math.min(img.width, img.height); const startX = (img.width - size) / 2; const startY = (img.height - size) / 2; canvas.width = 150; canvas.height = 150; ctx.clearRect(0, 0, 150, 150); ctx.drawImage(img, startX, startY, size, size, 0, 0, 150, 150); const base64Str = canvas.toDataURL('image/png'); document.getElementById(hiddenId).value = base64Str; let preview = document.getElementById(previewId); preview.src = base64Str; preview.classList.remove('hidden'); }; }; }

  function toggleSumberData() {
      let sumber = document.getElementById('pSumberData').value;
      if (sumber === 'import') document.getElementById('boxImportExcel').classList.remove('hidden');
      else document.getElementById('boxImportExcel').classList.add('hidden');
  }

  function formatTanggalExcel(excelDate) {
      if(!excelDate) return "";
      if(!isNaN(excelDate)) { let date = new Date((excelDate - (25567 + 2)) * 86400 * 1000); return date.toISOString().split('T')[0]; }
      let strDate = String(excelDate).trim(); let regexBaku = /^\d{4}-\d{2}-\d{2}$/;
      if(regexBaku.test(strDate)) return strDate;
      let parts = strDate.split(/[-/]/);
      if(parts.length === 3) {
          if(parts[0].length === 4) return `${parts[0]}-${parts[1].padStart(2,'0')}-${parts[2].padStart(2,'0')}`;
          return `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
      }
      return "1970-01-01";
  }

  function unduhTemplateExcelCustom(jenis) {
      document.getElementById('importJenis').value = jenis;
      unduhTemplateExcel();
  }

  // Paste ini di js_main.js
async function unduhTemplatePergub() {
    try {
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Template_Master_TPP');
        
        // Bikin Header
        sheet.addRow(["Nama Jabatan", "Kelas Jabatan", "BK", "PK", "KK", "TB", "KP", "Status Pegawai"]);
        
        // Kasih Contoh Data (Baris 2)
        sheet.addRow(["Kepala Dinas Pendidikan", "14", 5000000, 3000000, 0, 0, 0, "PNS"]);
        sheet.addRow(["Guru Ahli Pertama", "9", 1500000, 1000000, 0, 0, 0, "PPPK"]);
        
        // Warnain Header biar cantik
        sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        sheet.getRow(1).fill = { type: 'pattern', pattern:'solid', fgColor:{ argb:'FF2980B9' } };
        
        // Lebarkan kolom
        sheet.columns.forEach(column => { column.width = 20; });

        // Simpan & Unduh
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        
        // Buat link download palsu lalu klik otomatis
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = "Template_Import_Master_TPP.xlsx";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    } catch (error) {
        alertError("Gagal membuat template: " + error.message);
    }
}

  let loadingTimer;

  // ==============================================================
  // ALAT BANTU (HELPER) WAJIB FRONTEND - ANTI GHOST TIMER
  // ==============================================================
  let loadingInterval; // Variabel global untuk menyimpan timer

function startLoading(titleAwal) {
    let detik = 0;
    
    Swal.fire({
        title: titleAwal,
        html: `<div id="pesanLoading" class="mt-2 text-primary fw-bold">Menghubungkan ke server...</div>`,
        allowOutsideClick: false,
        showConfirmButton: false,
        didOpen: () => {
            Swal.showLoading();
            
            // Jalankan Timer Interaktif
            loadingInterval = setInterval(() => {
                detik++;
                let el = document.getElementById('pesanLoading');
                if (!el) return;

                // LOGIKA FLEKSIBEL: Teks berubah sesuai durasi tunggu
                if (detik === 2) {
                    el.innerText = "Membongkar arsip database...";
                    el.className = "mt-2 text-success fw-bold";
                } 
                else if (detik === 5) {
                    el.innerText = "Mohon tunggu sebentar lagi, server sedang sibuk...";
                    el.className = "mt-2 text-warning fw-bold";
                } 
                else if (detik === 10) {
                    el.innerText = "Sedikit lagi selesai, hampir siap...";
                    el.className = "mt-2 text-danger fw-bold";
                }
                else if (detik === 15) {
                    el.innerText = "Koneksi internet agak lambat, tetap menunggu...";
                }
            }, 1000); // Cek setiap 1 detik
        }
    });
}

function stopLoading() {
    // Hentikan timer interaktif agar tidak makan memori
    if (loadingInterval) clearInterval(loadingInterval);
    Swal.close();
}

  function alertSukses(pesan) { Swal.fire({ icon: 'success', title: 'Berhasil!', text: pesan, confirmButtonColor: '#198754' }); }
  function alertError(pesan) { Swal.fire({ icon: 'error', title: 'Oops...', text: pesan, confirmButtonColor: '#dc3545' }); }
  function alertPeringatan(pesan) { Swal.fire({ icon: 'warning', title: 'Perhatian', text: pesan, confirmButtonColor: '#ffc107' }); }

  window.addEventListener('scroll', function() { let btnScroll = document.getElementById('btnScrollTop'); if (window.scrollY > 200) { btnScroll.style.display = 'flex'; } else { btnScroll.style.display = 'none'; } });
  function toggleTombolPegawaiHP(buka) { let btnP = document.getElementById('mNavPegawai'); if(btnP) { if(buka) btnP.classList.remove('disabled'); else btnP.classList.add('disabled'); } }

  function getGroupPPPK(golDasar) {
      const g = String(golDasar).trim().toUpperCase();
      if(["XVII", "XVI", "XV", "XIV", "XIII"].includes(g)) return "XVII - XIII";
      if(["XII", "XI", "X", "IX"].includes(g)) return "XII - IX";
      if(["VIII", "VII", "VI", "V"].includes(g)) return "VIII - V";
      return "IV - I";
  }

  const numToLet = (n) => { let s=""; while(n>=0){ s=String.fromCharCode((n%26)+65)+s; n=Math.floor(n/26)-1; } return s; };

  function getSmartBulanAngka() {
      let reguler = [];
      if (arrayPeriode && arrayPeriode.length > 0) {
          reguler = arrayPeriode
              .filter(p => p.jenisPeriode === "Reguler" || !p.jenisPeriode)
              .map(p => parseInt(p.bulanAngka))
              .filter(n => !isNaN(n));
      }
      
      for (let i = 1; i <= 12; i++) {
          if (!reguler.includes(i)) return i; 
      }
      return 1; 
  }

  function setupAutocomplete(inp, arr) {
      if (!inp) return;
      let newInp = inp.cloneNode(true);
      inp.parentNode.replaceChild(newInp, inp);
      inp = newInp;

      inp.addEventListener("input", function(e) {
          let a, b, val = this.value;
          closeAllLists();
          if (!val) return false;
          
          a = document.createElement("DIV");
          a.setAttribute("id", this.id + "autocomplete-list");
          a.setAttribute("class", "autocomplete-items");
          this.parentNode.appendChild(a);
          
          let count = 0;
          for (let i = 0; i < arr.length; i++) {
              let displayStr = typeof arr[i] === 'object' ? arr[i].display : String(arr[i]);
              let valStr = typeof arr[i] === 'object' ? arr[i].value : String(arr[i]);

              if (displayStr.toLowerCase().includes(val.toLowerCase()) || valStr.toLowerCase().includes(val.toLowerCase())) {
                  count++;
                  b = document.createElement("DIV");
                  let regex = new RegExp(`(${val})`, "gi");
                  b.innerHTML = displayStr.replace(regex, "<strong class='text-primary'>$1</strong>");
                  b.innerHTML += "<input type='hidden' value='" + valStr.replace(/'/g, "&#39;") + "'>";
                  b.addEventListener("click", function(e) {
                      inp.value = this.getElementsByTagName("input")[0].value;
                      closeAllLists();
                  });
                  a.appendChild(b);
              }
              if(count >= 20) break; 
          }
      });
  }

  function closeAllLists(elmnt) {
      let x = document.getElementsByClassName("autocomplete-items");
      for (let i = 0; i < x.length; i++) {
          if (elmnt != x[i] && elmnt != document.getElementById("mUnitKerja") && elmnt != document.getElementById("mJabatan") && elmnt != document.getElementById("uUnitKerjaOp")) {
              x[i].parentNode.removeChild(x[i]);
          }
      }
  }
  
  document.addEventListener("click", function (e) { closeAllLists(e.target); });

  // TAMBAHKAN FUNGSI BARU INI DI MANA SAJA DI DALAM js_main.js
function cekBulanReferensi(namaBulanPilihan, idWarningBox) {
    let warningBox = document.getElementById(idWarningBox);
    if(!warningBox) return;
    
    // Cek apakah bulan yang dipilih sudah ada di database (arrayPeriode)
    let sudahAda = arrayPeriode.some(p => p.namaPeriode === namaBulanPilihan);
    
    // Jika belum ada, munculkan warning. Jika sudah ada, sembunyikan.
    if (!sudahAda && namaBulanPilihan !== "") {
        warningBox.classList.remove('hidden');
    } else {
        warningBox.classList.add('hidden');
    }
}

// FUNGSI BARU UNTUK IMPORT SUSULAN DENGAN PERINGATAN (SWEETALERT)
async function prosesImportUpdateExcel() {
    let fileInput = document.getElementById('eFileImportPegawai');
    if(!fileInput.files[0]) return alertPeringatan("Pilih file Excel SKP terlebih dahulu!");
    
    let bulanTarget = document.getElementById('eBulanNama').value;
    
    // Tampilkan Peringatan Konfirmasi sebelum memproses
    Swal.fire({
        title: 'Peringatan Import Susulan!',
        html: `Harap berhati-hati! Proses ini akan <b>MENIMPA</b> data:<br>
               <ul class="text-start mt-2 mb-2 text-danger fw-bold">
                 <li>Nilai SKP</li>
                 <li>Pangkat / Golongan</li>
                 <li>Unit Kerja & Unor Induk</li>
               </ul>
               untuk pegawai yang NIP-nya cocok di periode <b>${bulanTarget}</b>.<br><br>
               <span class="text-success"><i>(Data Gaji, Tunjangan, Status Kawin, dan Jabatan tetap AMAN).</i></span><br><br>
               Apakah Anda yakin ingin menimpa data tersebut?`,
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#dc3545', // Warna merah (Bahaya/Warning)
        cancelButtonColor: '#6c757d', // Warna abu-abu (Batal)
        confirmButtonText: '<i class="bi bi-exclamation-triangle"></i> Ya, Timpa Data!',
        cancelButtonText: 'Batal'
    }).then(async (result) => {
        // Jika user mengklik tombol "Ya, Timpa Data!"
        if (result.isConfirmed) {
            
            // Kita "tipu" sistem sebentar agar mengira ini import pegawai biasa
            let oldJenis = document.getElementById('importJenis').value;
            document.getElementById('importJenis').value = 'pegawai';

            // Panggil fungsi inti yang sudah ada (prosesImportExcel)
            await prosesImportExcel(fileInput, bulanTarget);

            // Kembalikan statusnya ke semula
            document.getElementById('importJenis').value = oldJenis;
            
            // Kosongkan file input agar bisa dipakai lagi nanti
            fileInput.value = "";

            // Tutup modal edit periode
            let modalObj = bootstrap.Modal.getInstance(document.getElementById('modalEditPeriode'));
            if(modalObj) modalObj.hide();
            
            // Refresh otomatis data pegawai di layar untuk melihat perubahannya
            muatDataPegawai(true); 
        }
    });
}
