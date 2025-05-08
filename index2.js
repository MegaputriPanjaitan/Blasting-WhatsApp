const venom = require('venom-bot');
const XLSX = require('xlsx');

// Fungsi delay (milidetik)
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Baca data dari Excel
const workbook = XLSX.readFile('dataTest.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet);

// Fungsi untuk mengubah format bulan ke Bahasa Indonesia
function formatBulanTahun(excelDateValue) {
  if (typeof excelDateValue === 'number') {
    const jsDate = new Date((excelDateValue - 25569) * 86400 * 1000);
    const bulan = jsDate.toLocaleString('id-ID', { month: 'long' });
    const tahun = jsDate.getFullYear();
    return `${bulan} ${tahun}`;
  }

  if (excelDateValue instanceof Date) {
    const bulan = excelDateValue.toLocaleString('id-ID', { month: 'long' });
    const tahun = excelDateValue.getFullYear();
    return `${bulan} ${tahun}`;
  }

  const parsedDate = new Date(excelDateValue);
  if (!isNaN(parsedDate)) {
    const bulan = parsedDate.toLocaleString('id-ID', { month: 'long' });
    const tahun = parsedDate.getFullYear();
    return `${bulan} ${tahun}`;
  }

  throw new Error('Format data bulan tidak dikenali.');
}

// Fungsi greeting otomatis berdasarkan waktu
function generateGreeting() {
  const hour = new Date().getHours();
  if (hour >= 5 && hour < 11) return 'Selamat pagi';
  if (hour >= 11 && hour < 15) return 'Selamat siang';
  if (hour >= 15 && hour < 18) return 'Selamat sore';
  return 'Selamat malam';
}

// Template pesan
function generateMessage(namaPerusahaan, bulan) {
  const greeting = generateGreeting();
  return `${greeting} Pak/Bu,

Perkenalkan, saya merupakan petugas admin dari Bapak Habieb selaku AR BPJS Ketenagakerjaan Cabang Medan Kota. 
Kami ingin mengonfirmasi apakah ini benar dengan pihak ${namaPerusahaan}?

Jika benar, kami ingin memastikan status pembayaran iuran BPJS Ketenagakerjaan pada bulan ${bulan}. Apabila pembayaran sudah dilakukan, mohon kirimkan bukti pembayarannya agar kami dapat melakukan pengecekan lebih lanjut.

Jika pembayaran belum dilakukan, kami informasikan bahwa sesuai arahan pimpinan, iuran wajib dibayarkan dalam bulan berjalan. Oleh karena itu, kami mengimbau Bapak/Ibu untuk segera menyelesaikan pembayaran.

Apakah Bapak/Ibu berencana untuk tetap melanjutkan pembayaran? Jika tidak, mohon konfirmasinya agar kami dapat memproses penonaktifan kepesertaan badan usaha Bapak/Ibu (sesuai dengan S&K yang berlaku).

Terima kasih atas perhatian dan kerja samanya.

Salam,
BPJS Ketenagakerjaan Cabang Medan Kota`;
}

// Mulai venom bot
venom
  .create({
    session: 'session-name',
    multidevice: true,
    headless: false,
  })
  .then(async (client) => {
    for (let i = 0; i < data.length; i++) {
      const { nama_perusahaan, bulan, nomor_whatsapp } = data[i];
      try {
        const formattedBulan = formatBulanTahun(bulan);
        const message = generateMessage(nama_perusahaan, formattedBulan);
        const nomor = `${nomor_whatsapp}@c.us`;

        await client.sendText(nomor, message);
        console.log(`✅ Pesan terkirim ke ${nama_perusahaan} (${nomor_whatsapp})`);

        await delay(5000); // delay 5 detik sebelum ke nomor berikutnya
      } catch (error) {
        console.error(`❌ Gagal kirim ke ${nama_perusahaan} (${nomor_whatsapp})`, error.message);
      }
    }
  })
  .catch((error) => {
    console.error('❌ Gagal membuat client venom:', error);
  });
