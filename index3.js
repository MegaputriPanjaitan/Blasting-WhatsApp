const venom = require('venom-bot');
const xlsx = require('xlsx');
const path = require('path');

// Baca file Excel
const workbook = xlsx.readFile('admin4.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet);

venom
  .create({
    session: 'session-wa-blast',
    headless: false, // tampilkan browser agar kamu bisa scan QR
    useChrome: true,
  })
  .then((client) => start(client))
  .catch((error) => {
    console.error(error);
  });


async function start(client) {
  for (const row of data) {
    const namaPerusahaan = row.nama_perusahaan;
    const tenagaKerja = row.tenaga_kerja;
    let nomor = row.nomor_whatsapp;

    // Format nomor (jika belum pakai @c.us)
    if (!nomor.endsWith('@c.us')) {
      nomor = nomor + '@c.us';
    }

    // Template pesan dari kamu, tidak diubah:
    const pesan = 
`Yth. Bapak/Ibu PIC Perusahaan ${namaPerusahaan} 
Dengan hormat,

Berdasarkan data yang kami miliki, terdapat tenaga kerja atas nama ${tenagaKerja} yang tercatat berusia ≥65 tahun. Sehubungan dengan hal tersebut, kami mohon kesediaan Bapak/Ibu untuk:

Melampirkan foto KTP dan foto terbaru dari tenaga kerja yang bersangkutan, serta membuat surat pernyataan yang menyatakan bahwa tenaga kerja tersebut masih aktif bekerja di perusahaan Bapak/Ibu.

Atas perhatian dan kerja sama Bapak/Ibu, kami ucapkan terima kasih.

Hormat kami,
Petugas Admin Bapak Habieb selaku AR BPJS KETENAGAKERJAAN Medan Kota`;

    try {
      // Kirim pesan
      await client.sendText(nomor, pesan);
      console.log(`✅ Pesan terkirim ke ${nomor}`);

      // Kirim lampiran surat
      const filePath = path.resolve(__dirname, 'surat_keterangan_aktif.pdf');
      await client.sendFile(nomor, filePath, 'surat.pdf', 'Berikut contoh suratnya Bapak/Ibu🙏');
      console.log(`📎 Surat terkirim ke ${nomor}`);
    } catch (error) {
      console.error(`❌ Gagal mengirim ke ${nomor}:`, error);
    }
  }
}
