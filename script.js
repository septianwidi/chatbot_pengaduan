// ====================== KONFIGURASI ======================
const telegramApiUrl = `https://api.telegram.org/bot${botToken}`;
const adminIds = getAdminChatIds();
const prop = PropertiesService.getScriptProperties();

// ====================== LOGGING ======================
function log(msg = '') {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(logSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(logSheetName);
    sheet.appendRow(['Waktu', 'Pesan']);
  }
  sheet.appendRow([new Date(), msg]);
}

// ====================== FORMAT TANGGAL ======================
function formatDate(date) {
  const month = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Ags','Sep','Okt','Nov','Des'];
  return `${date.getDate()} ${month[date.getMonth()]} ${date.getFullYear()}`;
}

// ====================== KIRIM PESAN TELEGRAM ======================
function sendTelegramMessage(chatId, replyToMessageId, text, keyboard = null) {
  const payload = {
    chat_id: chatId,
    parse_mode: 'HTML',
    text,
    disable_web_page_preview: true
  };
  if (replyToMessageId) payload.reply_to_message_id = replyToMessageId;
  if (keyboard) payload.reply_markup = keyboard;

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const res = UrlFetchApp.fetch(`${telegramApiUrl}/sendMessage`, options);
    log(`📤 Pesan dikirim ke ${chatId}: ${text} | resp:${res.getResponseCode()}`);
    return res.getContentText();
  } catch (err) {
    log(`❌ Gagal kirim pesan ke ${chatId}: ${err}`);
    return null;
  }
}

// ====================== TOKEN HELPERS ======================
function generateToken(len = 32) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let t = '';
  for (let i = 0; i < len; i++) t += chars.charAt(Math.floor(Math.random() * chars.length));
  return t;
}

function storeTokenForRow(row, token, expiryIso) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(dataPengaduanSheetName);
  sheet.getRange(row, 15).setValue(token);
  sheet.getRange(row, 16).setValue(expiryIso);
  sheet.getRange(row, 17).setValue('TIDAK');
}

function findRowByToken(token) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(dataPengaduanSheetName);
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const range = sheet.getRange(2, 15, last - 1, 1).getValues();
  for (let i = 0; i < range.length; i++) {
    if ((range[i][0] || '').toString() === token) return 2 + i;
  }
  return null;
}

function markTokenUsed(row) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  ss.getSheetByName(dataPengaduanSheetName).getRange(row, 17).setValue('YA');
}

function isTokenValid(row) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(dataPengaduanSheetName);
  const token = sheet.getRange(row, 15).getValue();
  const expiry = sheet.getRange(row, 16).getValue();
  const used = sheet.getRange(row, 17).getValue();
  if (!token) return false;
  if ((used || '').toString().toUpperCase() === 'YA') return false;
  if (!expiry) return false;
  const now = new Date();
  const exp = new Date(expiry);
  return now <= exp;
}

// ====================== AMBIL ADMIN DARI SHEET ======================
function getAdminChatIds() {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName('Admin');
  if (!sheet) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return data.map(r => r[0]).filter(v => v);
}

// ====================== MENU AWAL ======================
function menuAwal() {
  return {
    inline_keyboard: [
      [{ text: '📧 Pengaduan Email Kedinasan', callback_data: 'menu_layanan_email' }],
      [{ text: '🏢 Pengaduan eOffice', callback_data: 'menu_layanan_eoffice' }],
      [{ text: '📱 Pengaduan MyPKP Mobile', callback_data: 'menu_layanan_mypkp' }],
      [{ text: '🔍 Cek Status Tiket', callback_data: 'menu_tiket' }],
      [{ text: 'ℹ️ Panduan Format Pengaduan', callback_data: 'menu_format' }]
    ]
  };
}

// ====================== SUB MENU LAYANAN EMAIL======================
function layanan_email() {
  return {
    inline_keyboard: [
      [{ text: '📝 Buat Pengaduan Email Kedinasan', callback_data: 'menu_pengaduan_email' }],
      [{ text: '📚 FAQ Layanan Email Kedinasan', callback_data: 'menu_faq_email' }],
      [{ text: '👨‍💼 Hubungi Admin Email Kedinasan', url: 'https://t.me/septianwidi' }],
      [{ text: '⬅️ Menu Utama', callback_data: 'menu_home' }]
    ]
  };
}

// ====================== SUB MENU LAYANAN EOFFICE======================
function layanan_eoffice() {
  return {
    inline_keyboard: [
      [{ text: '📝 Buat Pengaduan eOffice', callback_data: 'menu_pengaduan_eoffice' }],
      [{ text: '📚 FAQ Layanan eOffice PKP', callback_data: 'menu_faq_eoffice' }],
      [{ text: '👨‍💼 Hubungi Admin eOffice', url: 'https://t.me/septianwidi' }],
      [{ text: '⬅️ Menu Utama', callback_data: 'menu_home' }]
    ]
  };
}

// ====================== SUB MENU LAYANAN MYPKP======================
function layanan_mypkp() {
  return {
    inline_keyboard: [
      [{ text: '📝 Buat Pengaduan MyPKP Mobile', callback_data: 'menu_pengaduan_mypkp' }],
      [{ text: '📚 FAQ Layanan MyPKP Mobile', callback_data: 'menu_faq_mypkp' }],
      [{ text: '👨‍💼 Hubungi Admin MyPKP Mobile', url: 'https://t.me/septianwidi' }],
      [{ text: '⬅️ Menu Utama', callback_data: 'menu_home' }]
    ]
  };
}

// submenu pengaduan layanan email
function menuPengaduanEmail() {
  return {
    inline_keyboard: [
      [{ text: '🔑 Reset Password Email Kedinasan', callback_data: 'menu_resetpw_email' }],
      [{ text: '📱 Reset MFA (Authenticator) Email Kedinasan', callback_data: 'menu_resetmfa_email' }],
      [{ text: '📝 Lainnya terkait Email Kedinasan', callback_data: 'menu_lainnya_email' }],
      [{ text: '⬅️ Menu Layanan Email Kedinasan', callback_data: 'menu_layanan_email' }]
    ]
  };
}

// submenu pengaduan layanan eoffice
function menuPengaduanEoffice() {
  return {
    inline_keyboard: [
      [{ text: '📨 Persuratan eOffice', callback_data: 'menu_persuratan' }],
      [{ text: '🏢 Mutasi Unit Kerja/Jabatan eOffice', callback_data: 'menu_mutasi' }],
      [{ text: '📝 Lainnya terkait eOffice', callback_data: 'menu_lainnya_eoffice' }],
      [{ text: '⬅️ Menu Layanan eOffice', callback_data: 'menu_layanan_eoffice' }]
    ]
  };
}

// submenu pengaduan layanan mypkp
function menuPengaduanMypkp() {
  return {
    inline_keyboard: [
      [{ text: '🔑 Reset Password MyPKP Mobile', callback_data: 'menu_resetpw_mypkp' }],
      [{ text: '📱 Kendala Presensi MyPKP Mobile', callback_data: 'menu_presensi_mypkp' }],
      [{ text: '📝 Lainnya terkait MyPKP Mobile', callback_data: 'menu_lainnya_mypkp' }],
      [{ text: '⬅️ Menu Layanan MyPKP Mobile', callback_data: 'menu_layanan_mypkp' }]
    ]
  };
}

// ====================== PARSER PESAN INPUT ======================
function parseMessage(message = '') {
  const lines = message.split('\n');
  const data = { nama:'', nip:'', email:'', no_hp:'', isi_aduan:'', status:'' };

  lines.forEach(line => {
    const [key, ...rest] = line.split(':');
    const val = rest.join(':').trim();
    if (!key) return;
    const k = key.trim().toLowerCase();

    if (k.startsWith('nama')) data.nama = val;
    else if (k.startsWith('nip')) data.nip = val;
    else if (k.startsWith('email')) data.email = val;
    else if (k.startsWith('no hp') || k.startsWith('no handphone') || k.startsWith('no. hp')) data.no_hp = val;
    else if (k.startsWith('deskripsi') || k.startsWith('uraian')) data.isi_aduan = val;
  });

  return Object.values(data).every(v => v === '') ? false : data;
}

// ====================== SIMPAN DATA ======================
  function inputDataPengaduan(data) {
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(dataPengaduanSheetName);
  const lastRow = sheet.getLastRow();
  const today = new Date();
  const dd = String(today.getDate()).padStart(2, '0');
  const monthNames = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  const mm = monthNames[today.getMonth()];
  const yy = String(today.getFullYear()).slice(-2);
  const tanggalFormat = `${dd}${mm}${yy}`;

  let nomorUrut = 1;
  if (lastRow > 1) {
    const lastTiket = sheet.getRange(lastRow, 9).getValue().toString();
    const match = lastTiket.match(/pkp-(\d{2}[A-Z]{3}\d{2})-(\d{3})/);
    if (match && match[1] === tanggalFormat) nomorUrut = parseInt(match[2], 10) + 1;
  }

  const noFormatted = String(nomorUrut).padStart(3, '0');
  const id_pengaduan = `ADU-${nomorUrut}`;
  const tiket = `pkp-${tanggalFormat}-${noFormatted}`;

  // Ambil kategori dari properti (diset dari menu pengaduan)
  const prop = PropertiesService.getScriptProperties();
  const kategori = prop.getProperty(`kategori_${data.chatId}`) || 'Lainnya';

  // Ambil Layanan dari properti (diset dari menu awal)
  const layanan = prop.getProperty(`layanan_${data.chatId}`);

  // Simpan data ke sheet sesuai urutan kolom
  sheet.appendRow([
    nomorUrut,                 // A
    id_pengaduan,              // B
    today,                     // C
    data.nama,                 // D
    data.nip,                  // E
    data.email,                // F
    data.no_hp,                // G
    data.isi_aduan,            // H
    tiket,                     // I
    kategori,                  // J
    'menunggu verifikasi admin', // K
    data.chatId,               // L
    '',                        // M PasswordBaru
    '',                        // N Catatan Admin
    '',                        // O Token
    '',                        // P Kadaluarsa
    '',                        // Q Sudah Digunakan
    '',                        // R Rating
    '',                         // S Feedback
    layanan                     // T Layanan
  ]);

  // 🧹 Hapus kategori agar tidak terbawa ke input berikutnya
  prop.deleteProperty(`kategori_${data.chatId}`);

  // 🧹 Hapus layanan agar tidak terbawa ke input berikutnya
  prop.deleteProperty(`layanan_${data.chatId}`);

  return tiket;
}

// ====================== CEK TIKET ======================
function cekAduan(tiket) {
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(dataPengaduanSheetName);
  const last = sheet.getLastRow();
  if (last < 2) return `⚠️ Tiket ${tiket} tidak ditemukan.`;
  const list = sheet.getRange(2, 1, last - 1, 17).getValues();
  const data = list.find(el => el[8] && el[8].toString().toLowerCase() === tiket.toLowerCase());
  if (!data) return `⚠️ Tiket ${tiket} tidak ditemukan.`;

  // Indeks kolom (0-based):
  // 0 => A (nomorUrut)
  // 1 => B (id_pengaduan)
  // 2 => C (tanggal)
  // 3 => D (nama)
  // 4 => E (nip)
  // 5 => F (email)
  // 6 => G (no_hp)
  // 7 => H (isi_aduan)
  // 8 => I (tiket)
  // 9 => J (kategori)
  // 10 => K (status) <-- benar
  const idPengaduan = data[1] || '-';
  const tanggalCell = data[2];
  const tanggalText = tanggalCell ? formatDate(new Date(tanggalCell)) : '-';
  const nama = data[3] || '-';
  const nip = data[4] || '-';
  const email = data[5] || '-';
  const noHp = data[6] || '-';
  const deskripsi = data[7] || '-';
  const status = data[10] || '-';

  return `📄 <b>Info Tiket ${tiket}</b>\n\n` +
         `<b>ID Pengaduan:</b> ${idPengaduan}\n` +
         `<b>Tanggal:</b> ${tanggalText}\n` +
         `<b>Nama:</b> ${nama}\n` +
         `<b>NIP:</b> ${nip}\n` +
         `<b>Email:</b> ${email}\n` +
         `<b>No. HP:</b> ${noHp}\n` +
         `<b>Deskripsi:</b> ${deskripsi}\n` +
         `<b>Status:</b> <b>${status}</b>`;
}

function processRatingCallback(callbackQuery) {
  try {
    const parts = callbackQuery.data.split('_');
    if (parts.length !== 3) return;
    const row = parseInt(parts[1], 10);
    const rating = parseInt(parts[2], 10);
    if (isNaN(row) || isNaN(rating)) return;

    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(dataPengaduanSheetName);
    sheet.getRange(row, 18).setValue(rating);

    let stars = '';
    for (let i = 1; i <= 5; i++) stars += i <= rating ? '⭐ ' : '☆ ';

    const chatId = callbackQuery.message.chat.id;
    const messageId = callbackQuery.message.message_id;

    // edit pesan bintang
    UrlFetchApp.fetch(`${telegramApiUrl}/editMessageText`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: chatId,
        message_id: messageId,
        text: `${stars}\n\nPenilaian yang Anda berikan (${rating}/5)! 🙏`,
        parse_mode: 'HTML'
      }),
      muteHttpExceptions: true
    });

    // Simpan status bahwa user boleh isi saran
    prop.setProperty(`awaiting_feedback_${chatId}`, row.toString());

    // kirim pesan permintaan saran
    sendTelegramMessage(
      chatId,
      null,
      '🗒️ <b>Terima kasih atas waktu dan kepercayaan Anda.</b>\nKami sangat menghargai apabila Anda berkenan menyampaikan pesan apresiasi, ungkapan kepuasan, ataupun saran dan masukan guna membantu peningkatan kualitas layanan kami.\n\nSilakan menuliskan tanggapan Anda dengan membalas pesan ini'
    );

    adminIds.forEach(id => sendTelegramMessage(id, null, `⭐ Pelapor di baris ${row} memberi rating ${rating}/5.`));
    log(`⭐ Rating ${rating} disimpan untuk baris ${row}, meminta saran pengguna`);
  } catch (err) {
    log(`❌ Error processRatingCallback: ${err}`);
  }
}

// ====================== HANDLE PESAN ======================
function doPost(e) {
  try {
    if (!e || !e.postData) return;
    const contents = JSON.parse(e.postData.contents);

    if (contents.callback_query) {
      const cq = contents.callback_query; 
      const chatId = cq.from.id;
      const messageId = cq.message ? cq.message.message_id : null;
      const data = cq.data ? cq.data.toString() : '';

      // --- menu handler ---
      if (data === 'menu_pengaduan_email') {
        sendTelegramMessage(chatId, messageId, 'Silakan pilih jenis pengaduan:', menuPengaduanEmail());
        return;
      }

      if (data === 'menu_pengaduan_eoffice') {
        sendTelegramMessage(chatId, messageId, 'Silakan pilih jenis pengaduan:', menuPengaduanEoffice());
        return;
      }

      if (data === 'menu_pengaduan_mypkp') {
        sendTelegramMessage(chatId, messageId, 'Silakan pilih jenis pengaduan:', menuPengaduanMypkp());
        return;
      }

      if (data === 'menu_resetpw_email') {
      PropertiesService.getScriptProperties().setProperty(`kategori_${chatId}`,'Reset Password Email Kedinasan');
      sendTelegramMessage(chatId, messageId,
      '📝 Silakan salin format dengan awalan kata <b>/input</b> dan isi pengaduan <b>Reset Password Email Kedinasan</b> berikut:\n\n' +
      '<pre>/input\nNama: \nNIP: \nEmail: @pkp.go.id\nNo HP: \nDeskripsi Pengaduan: Reset password akun Microsoft</pre>'
      );
      return;
      }

      if (data === 'menu_resetmfa_email') {
      PropertiesService.getScriptProperties().setProperty(`kategori_${chatId}`,'Reset MFA Email Kedinasan');
      sendTelegramMessage(chatId, messageId,
      '📝 Silakan salin format dengan awalan kata <b>/input</b> dan isi pengaduan <b>Reset MFA Email Kedinasan</b> berikut:\n\n' +
      '<pre>/input\nNama: \nNIP: \nEmail: @pkp.go.id\nNo HP: \nDeskripsi Pengaduan: Reset koneksi Microsoft Authenticator (MFA) Email Kedinasan</pre>'
      );
      return;
      }

      if (data === 'menu_lainnya_email') {
      PropertiesService.getScriptProperties().setProperty(`kategori_${chatId}`,'Lainnya Email Kedinasan');
      sendTelegramMessage(chatId, messageId,
      '📝 Silakan salin format dengan awalan kata <b>/input</b> dan isi pengaduan <b>Kendala Lainnya terkait Email Kedinasan</b> berikut:\n\n' +
      '<pre>/input\nNama: \nNIP: \nEmail: @pkp.go.id\nNo HP: \nDeskripsi Pengaduan: [Jelaskan kendala yang Anda hadapi terkait email kedinasan]</pre>'
      );
        return;
      }

      if (data === 'menu_persuratan') {
      PropertiesService.getScriptProperties().setProperty(`kategori_${chatId}`,'Persuratan EOffice');
      sendTelegramMessage(chatId, messageId,
      '📝 Silakan salin format dengan awalan kata <b>/input</b> dan isi pengaduan <b>Kendala Persuratan EOffice</b> berikut:\n\n' +
      '<pre>/input\nNama: \nNIP: \nEmail: @pkp.go.id\nNo HP: \nDeskripsi Pengaduan: Tidak bisa upload surat / gagal Dispo / gagal TTE</pre>'
      );
      return;
      }

      if (data === 'menu_mutasi') {
      PropertiesService.getScriptProperties().setProperty(`kategori_${chatId}`,'Mutasi EOffice');
      sendTelegramMessage(chatId, messageId,
      '📝 Silakan salin format dengan awalan kata <b>/input</b> dan isi pengaduan <b>Mutasi Jabatan/Unit Kerja pada aplikasi EOffice</b> berikut:\n\n' +
      '<pre>/input\nNama: \nNIP: \nEmail: @pkp.go.id\nNo HP: \nDeskripsi Pengaduan: Mutasi/pindah jabatan dari Unit Kerja (A: sebutkan) ke (B:sebutkan)</pre>'
      );
      return;
      }

      if (data === 'menu_lainnya_eoffice') {
      PropertiesService.getScriptProperties().setProperty(`kategori_${chatId}`,'Lainnya eOffice');
      sendTelegramMessage(chatId, messageId,
      '📝 Silakan salin format dengan awalan kata <b>/input</b> dan isi pengaduan <b>Kendala Lainnya terkait eOffice</b> berikut:\n\n' +
      '<pre>/input\nNama: \nNIP: \nEmail: @pkp.go.id\nNo HP: \nDeskripsi Pengaduan: [Jelaskan kendala yang Anda hadapi terkait aplikasi eOffice]</pre>'
      );
        return;
      }

      if (data === 'menu_resetpw_mypkp') {
      PropertiesService.getScriptProperties().setProperty(`kategori_${chatId}`,'Reset Password MyPKP Mobile');
      sendTelegramMessage(chatId, messageId,
      '📝 Silakan salin format dengan awalan kata <b>/input</b> dan isi pengaduan <b>Reset Password MyPKP Mobile</b> berikut:\n\n' +
      '<pre>/input\nNama: \nNIP: \nEmail: @pkp.go.id\nNo HP: \nDeskripsi Pengaduan: Reset password akun MyPKP Mobile</pre>'
      );
      return;
      }

      
      if (data === 'menu_presensi_mypkp') {
      PropertiesService.getScriptProperties().setProperty(`kategori_${chatId}`,'Kendala Presensi MyPKP Mobile');
      sendTelegramMessage(chatId, messageId,
      '📝 Silakan salin format dengan awalan kata <b>/input</b> dan isi pengaduan <b>Kendala Presensi MyPKP Mobile</b> berikut:\n\n' +
      '<pre>/input\nNama: \nNIP: \nEmail: @pkp.go.id\nNo HP: \nDeskripsi Pengaduan: Kendala Presensi di aplikasi MyPKP Mobile</pre>'
      );
      return;
      }

      if (data === 'menu_lainnya_mypkp') {
      PropertiesService.getScriptProperties().setProperty(`kategori_${chatId}`,'Lainnya MyPKP Mobile');
      sendTelegramMessage(chatId, messageId,
      '📝 Silakan salin format dengan awalan kata <b>/input</b> dan isi pengaduan <b>Kendala Lainnya terkait MyPKP Mobile</b> berikut:\n\n' +
      '<pre>/input\nNama: \nNIP: \nEmail: @pkp.go.id\nNo HP: \nDeskripsi Pengaduan: [Jelaskan kendala yang Anda hadapi terkait aplikasi MyPKP Mobile]</pre>'
      );
        return;
      }

      if (data === 'menu_tiket') {
        sendTelegramMessage(chatId, messageId, '🔍 Untuk cek status, ketik:\n\n<pre>/tiket pkp-.....</pre>');
        return;
      }

      if (data === 'menu_format') {
      sendTelegramMessage(
      chatId,
      messageId,
     '📝 <b>Panduan Format Laporan</b>\n\n' +
     'Gunakan format berikut saat membuat laporan:\n\n' +
     '<pre>/input\n' +
      'Nama: [Nama Lengkap]\n' +
      'NIP: [18 Digit NIP]\n' +
      'Email: [nama@pkp.go.id]\n' +
      'No HP: [08xx atau 62xx]\n' +
      'Deskripsi: [Jelaskan masalah secara detail]</pre>\n\n' +
 
      '<b>Tips:</b>\n' +
      '✅ Pastikan semua data terisi lengkap dan benar\n' +
      '✅ Email wajib menggunakan domain <code>@pkp.go.id</code>\n' +
      '✅ Deskripsikan kendala dengan jelas / detail\n'+
      '✅ Menghubungi admin apabila ingin menyertakan bukti terkait kendala yang dihadapi.'
      );
      return;
      }

      if (data === 'menu_faq_email') {
      const faqText =
      '📚 <b>FAQ - Pertanyaan yang Sering Diajukan terkait Email Kedinasan</b>\n\n' +

      '❓ <b>Bagaimana cara reset password akun Email Kedinasan @pkp.go.id di Microsoft 365?</b>\n' +
      '✅ Demi keamanan data dan privasi pengguna,kebijakan reset password email kedinasan bisa dilakukan dari petugas  Pusat Data dan Informasi sebagai Admin. Bapak/Ibu yang mempunyai kendala lupa password atau hal serupa dapat melakukan pengaduan melalui menu <b>Buat Pengaduan → Reset Password</b>.\n\n' +

      '❓ <b>Bagaimana jika saya tidak bisa masuk akun Microsoft 365 karena tidak muncul kode notifikasi Microsoft Authenticator (MFA) di Handphone?</b>\n' +
      '✅ Bapak/Ibu dapat memilih menu <b>Buat Pengaduan → Reset MFA</b> untuk permintaan pemutusan koneksi Microsoft Authenticator pada handphone dan akan diarahkan untuk registrasi ulang handphone sebagai sarana Multi Factor Authentication (MFA) lewat Microsoft Authenticator.\n\n' +
      
      '❓ <b>Berapa lama pengaduan saya diproses?</b>\n' +
      '✅ Pengaduan Bapak/Ibu diproses maksimal <b>1x24 jam hari kerja</b> oleh admin.\n\n' +

      '❓ <b>Apakah saya akan diberitahu jika laporan selesai?</b>\n' +
      '✅ Ya, Bapak/Ibu akan menerima notifikasi otomatis di Telegram.\n\n' +

      '❓ <b>Apakah data saya aman?</b>\n' +
      '✅ Ya, seluruh data tersimpan di Workspace instansi dan hanya admin berwenang yang dapat mengaksesnya.';

      sendTelegramMessage(chatId, messageId, faqText, {
      inline_keyboard: [
      [{ text: '⬅️ Kembali', callback_data: 'menu_layanan_email' }]
        ]
      });
      return;
      }

      if (data === 'menu_faq_eoffice') {
      const faqText =
      '📚 <b>FAQ - Pertanyaan yang Sering Diajukan Terkait eOffice </b>\n\n' +

      '❓ Apa itu e-Office PKP?\n'+
      '✅ e-Office PKP adalah aplikasi sistem persuratan elektronik berbasis web untuk mengelola surat masuk, surat keluar, disposisi, dan arsip secara digital di lingkungan Kementerian PKP. Proses paraf, tanda tangan elektronik, dan pelacakan surat dilakukan otomatis dan terdokumentasi.\n\n'+

      '❓ Bagaimana cara login ke e-Office PKP?\n'+
      '✅ Masuk ke laman https://eoffice.pkp.go.id/, isi username, password, dan captcha, lalu klik Login. Jika lupa password, pilih Reset Password, dan link akan dikirimkan ke email yang terdaftar.\n\n'+

      '❓ Bagaimana cara membuat surat keluar baru?\n'+
      '✅ Masuk menu Buat Surat (Keluar), pilih template surat (Undangan, Nota Dinas, Surat Dinas, Surat Tugas), isi alur surat dan penerima, lalu klik Generate Preview dan Simpan.\n\n'+

      '❓ Bagaimana cara menambahkan penerima surat?\n'+
      '✅ Pada langkah kedua pembuatan surat, pilih penerima Internal, Multiple, atau Eksternal, lalu klik OK untuk menyimpan daftar penerima.\n\n'+

      '❓ Bagaimana cara memberikan tanda tangan elektronik?\n'+
      '✅ Pada tab menu Tindak Lanjut Surat, pilih tipe tanda tangan, lalu klik Tandatangani dan Kirim. Surat otomatis tersimpan dengan tanda tangan digital dan terkirim ke Penerima.\n\n'+

      '❓ Bagaimana jika tanda tangan pimpinan tidak tersedia di e-office?\n'+
      '✅ Silahkan laporkan menggunakan layanan bot pusdatin.';

      sendTelegramMessage(chatId, messageId, faqText, {
      inline_keyboard: [
      [{ text: '⬅️ Kembali', callback_data: 'menu_layanan_eoffice' }]
        ]
      });
      return;
      }
 
      if (data === 'menu_faq_mypkp') {
      const faqText =
      '📚 <b>FAQ - Pertanyaan yang Sering Diajukan Terkait eOffice </b>\n\n' +

      '❓ Apa itu e-Presensi PKP?\n'+
      '✅ Aplikasi absensi digital MyPKP untuk mencatat kehadiran pegawai (WFO, WFH, izin, dinas), lengkap dengan pelacakan lokasi (GPS)\n\n'+

      '❓ Bagaimana cara melakukan Presensi?\n'+
      '✅ Jika menggunakan browser, Akses https://eoffice.pkp.go.id/ dan login menggunakan akun MyPKP kemudian pilih presensi. Apabila mengakses melalui smartphone, pastikan aplikasi MyPKP telah di unduh. Pastikan GPS aktif dan browser memiliki izin lokasi.\n\n'+

      '❓ Kenapa presensi saya gagal?\n'+
      '✅ Pastikan GPS aktif, lokasi kantor terdeteksi, dan jaringan stabil. Jika masih gagal, gunakan /lapor presensi [uraian masalah].\n\n'+

      '❓ Bagaimana cara untuk mengajukan cuti tahunan?\n'+
      '✅ Pegawai dengan masa kerja >= 1 tahun bisa mengajukan cuti tahunan. Silahkan pilih cuti pengajuan kemudian tambah pengajuan kemudian pilih jenis cuti. Jika tidak tersedia jenis cuti. Silahkan melakukan pelaporan melalui chatbot Layanan Pusdatin.\n\n'+

      '❓ Bagaimana mekanisme melakukan cuti sakit?\n'+
      '✅ Silahkan pilih cuti pengajuan kemudian tambah pengajuan kemudian pilih jenis cuti. Jika tidak tersedia jenis cuti. Silahkan melakukan pelaporan melalui chatbot Layanan Pusdatin';

      sendTelegramMessage(chatId, messageId, faqText, {
      inline_keyboard: [
      [{ text: '⬅️ Kembali', callback_data: 'menu_layanan_mypkp' }]
        ]
      });
      return;
      }

      if (data === 'menu_home') {
        sendTelegramMessage(chatId, messageId, '🏠 Menu utama.', menuAwal());
        return;
      }

      if (data === 'menu_layanan_email') {
        sendTelegramMessage(chatId, messageId, '📧 Menu Layanan Email Kedinasan.', layanan_email());
        return;
      }

      if (data === 'menu_layanan_eoffice') {
        sendTelegramMessage(chatId, messageId, '🏢 Menu Layanan eOffice.', layanan_eoffice());
        return;
      }

      if (data === 'menu_layanan_mypkp') {
        sendTelegramMessage(chatId, messageId, '📱 Menu Layanan MyPKP Mobile.', layanan_mypkp());
        return;
      }

      if (data.startsWith('cek_')) {
      const tiket = data.replace('cek_', '');
      sendTelegramMessage(chatId, messageId, cekAduan(tiket));
      return;
      }

      if (data.startsWith('rate_')) {
        processRatingCallback(cq);
        return;
      }

      sendTelegramMessage(chatId, messageId, 'Perintah tidak dikenali. awali perintah dengan tanda "/input" ', menuAwal());
      return;
    }

    // --- pesan biasa ---
    const msg = contents.message;
    if (!msg) return;
    const chatId = msg.chat.id;
    const messageId = msg.message_id;
    const text = (msg.text || '').trim();
    const textLower = text.toLowerCase();

    // --- CEK apakah user sedang diminta memberikan saran ---
    const awaitingKey = `awaiting_feedback_${chatId}`;
    const awaitingRow = prop.getProperty(awaitingKey);

    if (awaitingRow) {
    const rowNum = parseInt(awaitingRow, 10);
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(dataPengaduanSheetName);
    const existingFeedback = sheet.getRange(rowNum, 19).getValue();

    // Jika sudah ada saran sebelumnya, tolak input baru
    if (existingFeedback && existingFeedback.toString().trim() !== '') {
    sendTelegramMessage(chatId, messageId, 'ℹ️ Terima kasih, namun saran untuk tiket ini sudah pernah dikirim sebelumnya. 🙏');
    prop.deleteProperty(awaitingKey);
    return;
    }

    // Simpan saran ke kolom 19 (S)
    sheet.getRange(rowNum, 19).setValue(text);

    // Kirim ucapan terima kasih TANPA menampilkan menu utama
    sendTelegramMessage(chatId, messageId, '🙏 Terima kasih atas saran dan masukan yang telah diberikan. Kami akan terus berupaya meningkatkan kualitas layanan ke depan. 💙 \n\n'+'<b>Tim Layanan Sistem Informasi</b>\n'+'<b>Pusat Data dan Informasi</b>\n'+'<b>Kementerian Perumahan dan Kawasan Permukiman</b>');

    // Kirim notifikasi ke admin
    adminIds.forEach(id =>
    sendTelegramMessage(id, null, `💬 Pelapor di baris ${rowNum} mengirim saran:\n"${text}"`)
    );

  // Catat log & hapus status "menunggu saran"
  log(`💬 Saran disimpan untuk baris ${rowNum}: ${text}`);
  prop.deleteProperty(awaitingKey);
  return;
  }

    if (textLower === '/start') {
  const greeting =
    '👋 Selamat datang.\n\n' +
    'Anda mengakses <b>Sistem Bot Layanan Pusat Data dan Informasi</b>\n' +
    'Kementerian Perumahan dan Kawasan Permukiman.\n\n' +
    'Layanan ini disediakan untuk memfasilitasi penyampaian dan pemantauan pengaduan terkait kendala penggunaan email kedinasan, kendala penggunaan eOffice dan kendala penggunaan MyPKP Mobile.\n' +
    'Silakan memilih menu yang tersedia di bawah ini ⤵️';

  sendTelegramMessage(chatId, messageId, greeting, menuAwal());
  return;
}


    if (textLower.startsWith('/input')) {
      const parsed = parseMessage(msg.text || '');
      if (!parsed) {
        sendTelegramMessage(chatId, messageId, '⚠️ Format salah. Ketik <b>/format</b> untuk contoh.');
        return;
      }

      parsed.chatId = chatId;
      const tiket = inputDataPengaduan(parsed);
      const keyboard = {
        inline_keyboard: [
          [{ text: '📄 Lihat Status', callback_data: `cek_${tiket}` }],
          [{ text: '🏠 Kembali ke Menu Utama', callback_data: 'menu_home' }]
        ]
      };

      sendTelegramMessage(chatId, messageId, `✅ Pengaduan tersimpan!\nNomor Tiket: <b>${tiket}</b>`, keyboard);

      const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
      const notif = 
      `📢 <b>Aduan Baru Masuk!</b>\n` +
      `<b>Nama:</b> ${parsed.nama}\n` +
      `<b>NIP:</b> ${parsed.nip}\n` +
      `<b>Email:</b> ${parsed.email}\n` +
      `<b>No HP:</b> ${parsed.no_hp}\n` +
      `<b>Deskripsi:</b> ${parsed.isi_aduan}\n` +
      `<b>Tiket:</b> <code>${tiket}</code>`;

      const keyboardAdmin = {
      inline_keyboard: [
      [
        { text: '📄 Lihat di Spreadsheet', url: spreadsheetUrl }
      ]
    ]
};

adminIds.forEach(id => sendTelegramMessage(id, null, notif, keyboardAdmin));
return;

    }

    if (textLower.startsWith('/tiket')) {
      const tiket = text.split(' ')[1] || '';
      if (!tiket) {
        sendTelegramMessage(chatId, messageId, '⚠️ Ketik nomor tiket, contoh: /tiket pkp-010126-001');
        return;
      }
      sendTelegramMessage(chatId, messageId, cekAduan(tiket));
      return;
    }

    sendTelegramMessage(chatId, messageId, 'Perintah tidak dikenali.', menuAwal());
  } catch (err) {
    log(`❌ Error doPost: ${err}`);
  }
}

// ====================== TRIGGER onEdit ======================
function onEdit(e) {
  try {
    if (!e || !e.range) {
      log("⚠️ Fungsi onEdit dijalankan tanpa event (manual run)");
      return;
    }

    const range = e.range;
    const sheet = range.getSheet();
    if (!sheet || sheet.getName() !== dataPengaduanSheetName) return;

    const row = range.getRow();
    const col = range.getColumn();

    const statusCol = 11; // K
    const chatIdCol = 12; // L
    const passwordBaruCol = 13; // M
    const usedCol = 17; // O


    if (col !== statusCol || row <= 1) return;

    const statusVal = String(sheet.getRange(row, statusCol).getValue()).trim().toLowerCase();
    const chatId = sheet.getRange(row, chatIdCol).getValue();
    const passwordBaru = sheet.getRange(row, passwordBaruCol).getValue();

    if (!chatId) {
      log(`⚠️ Tidak ada chatId di baris ${row}`);
      return;
    }

  // CASE 1: STATUS = "selesaipw_email dan selesaipw_mypkp" (Reset password, kirim link aman)
  if (statusVal === 'selesaipw_email' || statusVal === 'selesaipw_mypkp') {
  const alreadyUsed = sheet.getRange(row, usedCol).getValue();
  if ((alreadyUsed || '').toString().toUpperCase() === 'YA') {
    log(`ℹ️ Token sudah dipakai / notifikasi sudah dikirim (baris ${row})`);
    return;
  }

  // Ambil catatan admin dari kolom 14
  const catatanAdmin = sheet.getRange(row, 14).getValue() || '';

  const token = generateToken(48);
  const expiry = new Date(Date.now() + 24 * 60 * 60 * 1000); // 24 jam
  storeTokenForRow(row, token, expiry.toISOString());

  const safeUrl = `${appsScriptUrl}?token=${encodeURIComponent(token)}`;

  // Pesan ke pelapor
  let pesan =
    `✅ Laporan Anda telah diselesaikan.\n\n` +
    `🔐 Untuk melihat password baru secara aman, klik tautan berikut ` +
    `(kadaluarsa dalam 24 jam & hanya dapat dilihat 1x):\n\n`;

  // Tambahkan catatan admin bila ada
  if (catatanAdmin.trim() !== '') {
    pesan += `🗒️ <b>Catatan dari Admin:</b>\n${catatanAdmin}\n\n`;
  }

  const keyboard = { 
    inline_keyboard: [
      [{ text: '🔐 Lihat Password (Aman)', url: safeUrl }]
    ]
  };

  // Kirim link aman ke pelapor
  sendTelegramMessage(chatId, null, pesan, keyboard);

  // Permintaan rating
  const rateKeyboard = {
    inline_keyboard: [
      [
        { text: '☆', callback_data: `rate_${row}_1` },
        { text: '☆', callback_data: `rate_${row}_2` },
        { text: '☆', callback_data: `rate_${row}_3` },
        { text: '☆', callback_data: `rate_${row}_4` },
        { text: '☆', callback_data: `rate_${row}_5` }
      ]
    ]
  };

  sendTelegramMessage(
    chatId,
    null,
    '✨ Mohon berikan penilaian Anda terhadap kemudahan & keefektifan pelaporan melalui chatbot ini:',
    rateKeyboard
  );

  // Log & notifikasi admin
  const pesanAdmin = 
    `📬 Laporan baris ${row} diselesaikan (reset password).\n` +
    `Token: ${token.substring(0,8)}...\n` +
    `Link dikirim ke pelapor (${chatId}).`;

  adminIds.forEach(id => sendTelegramMessage(id, null, pesanAdmin));

  log(`🔗 Token dibuat & link + catatan admin dikirim ke pelapor (row ${row})`);
  return;
}

    // CASE 2: STATUS = "selesaimfa_email" (Putus koneksi MFA tanpa password)
    if (statusVal === 'selesaimfa_email') {
    sheet.getRange(row, usedCol).setValue('YA');

    // Ambil catatan admin dari kolom 14
    const catatanAdmin = sheet.getRange(row, 14).getValue() || '';

    // Pesan utama
    let pesan =
    `✅ Laporan Anda telah diselesaikan.\n\n` +
    `📱 Pemutusan koneksi Microsoft Authenticator telah dilakukan. ` +
    `Silakan tambahkan ulang akun Authenticator Anda sebelum login berikutnya.\n\n`;

    // Sisipkan catatan admin bila ada
    if (catatanAdmin.trim() !== '') {
    pesan += `🗒️ <b>Catatan dari Admin:</b>\n${catatanAdmin}\n\n`;
    }

    pesan += `Jika masih mengalami kendala, silakan hubungi admin.`;

    // Kirim ke pelapor
    sendTelegramMessage(chatId, null, pesan);

    // Notifikasi ke admin
    const pesanAdmin =
    `📬 Laporan baris ${row} diselesaikan (pemutusan MFA). ` +
    `Pesan konfirmasi dikirim ke pelapor (${chatId}).`;

    adminIds.forEach(id => sendTelegramMessage(id, null, pesanAdmin));

    // Permintaan rating
    const rateKeyboard = {
    inline_keyboard: [
      [
        { text: '☆', callback_data: `rate_${row}_1` },
        { text: '☆', callback_data: `rate_${row}_2` },
        { text: '☆', callback_data: `rate_${row}_3` },
        { text: '☆', callback_data: `rate_${row}_4` },
        { text: '☆', callback_data: `rate_${row}_5` }
      ]
    ]
  };

  sendTelegramMessage(
    chatId,
    null,
    '✨ Mohon berikan penilaian Anda terhadap kemudahan & keefektifan pelaporan melalui chatbot ini:',
    rateKeyboard
  );

  log(`ℹ️ Laporan MFA selesai & pesan dikirim ke pelapor (row ${row})`);
  return;
  }

  // CASE 3: STATUS = "selesaisurat_eoffice" (kendala surat eoffice sudah selesai)
    if (statusVal === 'selesaisurat_eoffice') {
    sheet.getRange(row, usedCol).setValue('YA');

    // Ambil catatan admin dari kolom 14
    const catatanAdmin = sheet.getRange(row, 14).getValue() || '';

    // Pesan utama
    let pesan =
    `✅ Laporan Anda telah diselesaikan.\n\n` +
    `📨 Kendala Persuratan eOffice yang Anda laporkan telah ditangani oleh admin. ` +
    `Silakan lakukan pengecekan kembali pada aplikasi eOffice.\n\n`;

    // Sisipkan catatan admin bila ada
    if (catatanAdmin.trim() !== '') {
    pesan += `🗒️ <b>Catatan dari Admin:</b>\n${catatanAdmin}\n\n`;
    }

    pesan += `Jika masih mengalami kendala, silakan hubungi admin.`;

    // Kirim ke pelapor
    sendTelegramMessage(chatId, null, pesan);

    // Notifikasi ke admin
    const pesanAdmin =
    `📬 Laporan baris ${row} diselesaikan (kendala persuratan eOffice). ` +
    `Pesan konfirmasi dikirim ke pelapor (${chatId}).`;

    adminIds.forEach(id => sendTelegramMessage(id, null, pesanAdmin));

    // Permintaan rating
    const rateKeyboard = {
    inline_keyboard: [
      [
        { text: '☆', callback_data: `rate_${row}_1` },
        { text: '☆', callback_data: `rate_${row}_2` },
        { text: '☆', callback_data: `rate_${row}_3` },
        { text: '☆', callback_data: `rate_${row}_4` },
        { text: '☆', callback_data: `rate_${row}_5` }
      ]
    ]
  };

  sendTelegramMessage(
    chatId,
    null,
    '✨ Mohon berikan penilaian Anda terhadap kemudahan & keefektifan pelaporan melalui chatbot ini:',
    rateKeyboard
  );

  log(`ℹ️ Laporan Persuratan eOffice selesai & pesan dikirim ke pelapor (row ${row})`);
  return;
  }

  // CASE 4: STATUS = "selesaimutasi_eoffice" (kendala surat eoffice sudah selesai)
    if (statusVal === 'selesaimutasi_eoffice') {
    sheet.getRange(row, usedCol).setValue('YA');

    // Ambil catatan admin dari kolom 14
    const catatanAdmin = sheet.getRange(row, 14).getValue() || '';

    // Pesan utama
    let pesan =
    `✅ Laporan Anda telah diselesaikan.\n\n` +
    `🏢 Jabatan dan Unit Kerja pada eOffice telah diperbaharui oleh admin. ` +
    `Silakan lakukan pengecekan kembali pada aplikasi eOffice.\n\n`;

    // Sisipkan catatan admin bila ada
    if (catatanAdmin.trim() !== '') {
    pesan += `🗒️ <b>Catatan dari Admin:</b>\n${catatanAdmin}\n\n`;
    }

    pesan += `Jika masih mengalami kendala, silakan hubungi admin.`;

    // Kirim ke pelapor
    sendTelegramMessage(chatId, null, pesan);

    // Notifikasi ke admin
    const pesanAdmin =
    `📬 Laporan baris ${row} diselesaikan (mutasi jabatan dan unitkerja pada eOffice). ` +
    `Pesan konfirmasi dikirim ke pelapor (${chatId}).`;

    adminIds.forEach(id => sendTelegramMessage(id, null, pesanAdmin));

    // Permintaan rating
    const rateKeyboard = {
    inline_keyboard: [
      [
        { text: '☆', callback_data: `rate_${row}_1` },
        { text: '☆', callback_data: `rate_${row}_2` },
        { text: '☆', callback_data: `rate_${row}_3` },
        { text: '☆', callback_data: `rate_${row}_4` },
        { text: '☆', callback_data: `rate_${row}_5` }
      ]
    ]
  };

  sendTelegramMessage(
    chatId,
    null,
    '✨ Mohon berikan penilaian Anda terhadap kemudahan & keefektifan pelaporan melalui chatbot ini:',
    rateKeyboard
  );

  log(`ℹ️ Laporan Mutasi Jabatan dan Unit Kerja pada eOffice selesai & pesan dikirim ke pelapor (row ${row})`);
  return;
  }


  // CASE 5: STATUS = "selesaipresensi_mypkp" (kendala presensi sudah selesai)
    if (statusVal === 'selesaipresensi_mypkp') {
    sheet.getRange(row, usedCol).setValue('YA');

    // Ambil catatan admin dari kolom 14
    const catatanAdmin = sheet.getRange(row, 14).getValue() || '';

    // Pesan utama
    let pesan =
    `✅ Laporan Anda telah diselesaikan.\n\n` +
    `📱 Kendala Presensi pada aplikasi MyPKP Mobile telah ditangani oleh admin. ` +
    `Silakan lakukan pengecekan kembali pada aplikasi MyPKP Mobile.\n\n`;

    // Sisipkan catatan admin bila ada
    if (catatanAdmin.trim() !== '') {
    pesan += `🗒️ <b>Catatan dari Admin:</b>\n${catatanAdmin}\n\n`;
    }

    pesan += `Jika masih mengalami kendala, silakan hubungi admin.`;

    // Kirim ke pelapor
    sendTelegramMessage(chatId, null, pesan);

    // Notifikasi ke admin
    const pesanAdmin =
    `📬 Laporan baris ${row} diselesaikan (presensi MyPKP Mobile). ` +
    `Pesan konfirmasi dikirim ke pelapor (${chatId}).`;

    adminIds.forEach(id => sendTelegramMessage(id, null, pesanAdmin));

    // Permintaan rating
    const rateKeyboard = {
    inline_keyboard: [
      [
        { text: '☆', callback_data: `rate_${row}_1` },
        { text: '☆', callback_data: `rate_${row}_2` },
        { text: '☆', callback_data: `rate_${row}_3` },
        { text: '☆', callback_data: `rate_${row}_4` },
        { text: '☆', callback_data: `rate_${row}_5` }
      ]
    ]
  };

  sendTelegramMessage(
    chatId,
    null,
    '✨ Mohon berikan penilaian Anda terhadap kemudahan & keefektifan pelaporan melalui chatbot ini:',
    rateKeyboard
  );

  log(`ℹ️ Laporan Kendala Presensi MyPKP Mobile selesai & pesan dikirim ke pelapor (row ${row})`);
  return;
  }

    // CASE 6: STATUS = "selesailainnya" (Pengaduan lainnya selesai)
    if (statusVal === 'selesailainnya_email' || statusVal === 'selesailainnya_eoffice' || statusVal === 'selesailainnya_mypkp') {
    sheet.getRange(row, usedCol).setValue('YA');

    // Ambil data dari sheet kolom 8 isiaduan dan kolom 14 catatan admin.
    const uraian = sheet.getRange(row, 8).getValue() || '(tidak ada uraian)';
    const catatanAdmin = sheet.getRange(row, 14).getValue() || ''; // kolom catatan admin 

    // Susun pesan ke pelapor
    let pesan =
    `✅ Laporan Anda telah diselesaikan.\n\n` +
    `📄 <b>Ringkasan Pengaduan Anda:</b>\n"${uraian}"\n\n` +
    `📨 Terima kasih telah melaporkan kendala Anda melalui sistem bot <b>Layanan PKP</b>. ` +
    `Permasalahan Anda telah kami tindak lanjuti dan dinyatakan <b>selesai</b>.`;

    // Tambahkan catatan admin bila ada
    if (catatanAdmin.trim() !== '') {
    pesan += `\n\n🗒️ <b>Catatan dari Admin:</b>\n${catatanAdmin}`;
    }

    pesan += `\n\nApabila masih ada kendala serupa, Anda dapat membuat pengaduan baru.`;

    // Kirim pesan ke pelapor
    sendTelegramMessage(chatId, null, pesan);

    // Kirim notifikasi ke admin
    const pesanAdmin = `📬 Laporan baris ${row} diselesaikan (pengaduan lainnya). Pesan konfirmasi dikirim ke pelapor (${chatId}).`;
    adminIds.forEach(id => sendTelegramMessage(id, null, pesanAdmin));

    // Kirim permintaan rating
    const rateKeyboard = {
    inline_keyboard: [
      [
        { text: '☆', callback_data: `rate_${row}_1` },
        { text: '☆', callback_data: `rate_${row}_2` },
        { text: '☆', callback_data: `rate_${row}_3` },
        { text: '☆', callback_data: `rate_${row}_4` },
        { text: '☆', callback_data: `rate_${row}_5` }
      ]
      ]
    };

    sendTelegramMessage(
    chatId,
    null,
    '✨ Mohon berikan penilaian Anda terhadap kemudahan & keefektifan layanan pengaduan ini:',
    rateKeyboard
  );

  log(`ℹ️ Laporan pengaduan lainnya selesai & pesan dikirim ke pelapor (row ${row})`);
  return;
  }

    // CASE 4: STATUS = "sedang diproses" (notifikasi proses laporan)
    if (statusVal.toString().trim().toLowerCase() === 'sedang diproses') {

    const pesan =
    `✅ Laporan Anda telah berhasil diverifikasi oleh admin.\n\n` +
    `🛠️ Saat ini pengaduan Anda sedang dalam proses penanganan oleh tim terkait.\n` +
    `Mohon menunggu, Anda akan mendapatkan update kembali setelah proses selesai.`;

    sendTelegramMessage(chatId, null, pesan);

    const pesanAdmin =
    `📬 Laporan baris ${row} telah berubah ke status *sedang diproses*.\n` +
    `Notifikasi pemberitahuan sudah dikirim ke pelapor (${chatId}).`;

    adminIds.forEach(id => sendTelegramMessage(id, null, pesanAdmin));

    log(`ℹ️ Status 'sedang diproses' — notifikasi dikirim ke pelapor (row ${row})`);
    return;
    }

  // CASE 5: STATUS = "data tidak valid" (Pengaduan ditolak/invalid)
  // Ambil data tersimpan dari sheet
  const nama   = sheet.getRange(row, 4).getValue() || '';
  const nip    = sheet.getRange(row, 5).getValue() || '';
  const email  = sheet.getRange(row, 6).getValue() || '';
  const noHp   = sheet.getRange(row, 7).getValue() || '';
  const aduan  = sheet.getRange(row, 8).getValue() || '';

  if (statusVal === 'data tidak valid') {
  sheet.getRange(row, usedCol).setValue('YA');

  // Ambil catatan admin dari kolom 14
  const catatanAdmin = sheet.getRange(row, 14).getValue() || '';

  // Pesan utama ke pelapor
  let pesan =
  `⚠️ <b>Data Pengaduan Tidak Valid</b>\n\n` +
  `Pengaduan Anda belum dapat diproses. Silakan lakukan perbaikan dengan menyesuaikan data berikut:\n\n` +
  `<pre>` +
  `/input\n` +
  `Nama: ${nama}\n` +
  `NIP: ${nip}\n` +
  `Email: ${email}\n` +
  `No HP: ${noHp}\n` +
  `Deskripsi Pengaduan: ${aduan}\n` +
  `</pre>\n\n`;

  // Jika ada catatan admin, tampilkan
  if (catatanAdmin.trim() !== '') {
    pesan += `🗒️ <b>Catatan dari Admin:</b>\n${catatanAdmin}\n\n`;
  }

  pesan += `Silakan melakukan perbaikan sesuai dengan catatan admin dan ajukan ulang pengaduan dengan informasi yang lengkap dan benar.\n` +
  `Terima kasih atas pengertiannya.`;

  // Kirim ke pelapor
  sendTelegramMessage(chatId, null, pesan);

  // Notifikasi admin
  const pesanAdmin =
    `📬 Laporan baris ${row} ditandai sebagai *data tidak valid*.\n` +
    `Catatan admin telah dikirim ke pelapor (${chatId}).`;

  adminIds.forEach(id => sendTelegramMessage(id, null, pesanAdmin));

  log(`ℹ️ Laporan 'data tidak valid' — notifikasi + catatan admin dikirim (row ${row})`);
  return;
}

  } catch (err) {
    log(`❌ Error onEdit: ${err}`);
  }
}

function doGet(e) {
  try {
    const token = (e.parameter && e.parameter.token) ? e.parameter.token.toString() : null;
    if (!token) {
      return HtmlService.createHtmlOutput('<p>Token tidak ditemukan.</p>');
    }

    const row = findRowByToken(token);
    if (!row) {
      return HtmlService.createHtmlOutput('<p>Token tidak valid atau sudah digunakan.</p>');
    }

    if (!isTokenValid(row)) {
      return HtmlService.createHtmlOutput('<p>Token telah digunakan atau kadaluarsa.</p>');
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(dataPengaduanSheetName);
    const passwordBaru = sheet.getRange(row, 13).getValue() || 'Password tidak tersedia!!';

    markTokenUsed(row);

    // 🔹 Gunakan link logo dari Google Drive (public)
    const logoUrl = 'https://lh3.googleusercontent.com/d/1RH2cTBNImCHUDmytrqiYwtn2kmaA5obk=s220';

    const html = `
      <html>
        <head>
          <meta name="viewport" content="width=device-width,initial-scale=1">
          <style>
            body {
              font-family: Arial, Helvetica, sans-serif;
              padding: 20px;
              color: #222;
              background: #f2f6f7;
              text-align: center;
            }
            .logo {
              display: block;
              margin: 20px auto 10px auto;
              max-width: 120px;
              height: auto;
            }
            .instansi {
              font-weight: bold;
              color: #87fff;
              font-size: 16px;
              margin-bottom: 25px;
            }
            .card {
              max-width: 600px;
              margin: 0 auto;
              padding: 25px 20px;
              border-radius: 10px;
              box-shadow: 0 4px 14px rgba(0,0,0,0.08);
              background-color: #0e5b73;
              color: #ffffff;
            }
            h2 {
              margin-top: 0;
              color: #ffffff;
            }
            .pw {
              font-size: 18px;
              padding: 12px;
              background: #ffffff;
              border-radius: 6px;
              word-break: break-all;
              color: #367588;
              font-weight: bold;
              display: inline-block;
              min-width: 220px;
            }
            .note {
              color: #ffffff;
              font-size: 13px;
              margin-top: 10px;
            }
          </style>
        </head>
        <body>
          <!-- Logo dan nama instansi di luar kotak -->
          <img src="${logoUrl}" alt="Logo Instansi" class="logo">
          
          <!-- Kotak password -->
          <div class="card">
            <h2>Link Password Sekali Pakai</h2>
            <p>Berikut password baru Anda (hanya dapat dilihat satu kali):</p>
            <div class="pw">${escapeHtml(passwordBaru)}</div>
            <p class="note">
              Pastikan untuk mengubah password setelah berhasil login.<br>
              Jika butuh bantuan, silakan hubungi admin.
            </p>
          </div>
        </body>
      </html>
    `;

    return HtmlService.createHtmlOutput(html);
  } catch (err) {
    log(`❌ Error doGet: ${err}`);
    return HtmlService.createHtmlOutput('<p>Terjadi kesalahan pada server.</p>');
  }
}

// helper escape html
function escapeHtml(text) {
  if (text === null || text === undefined) return '';
  return text.toString()
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#039;');
}

// ====================== SET WEBHOOK ======================
function setWebhook() {
  const res = UrlFetchApp.fetch(`${telegramApiUrl}/setWebhook?url=${appsScriptUrl}`).getContentText();
  Logger.log(res);
}

// ====================== REKAP RATING OTOMATIS ======================
function rekapRating() {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(dataPengaduanSheetName);
    const data = sheet.getDataRange().getValues();

    let totalRating = 0;
    let count = 0;

    // kolom 18 = Rating dihitung dari INDEX mulai dari 0
    for (let i = 1; i < data.length; i++) {
      const rating = parseFloat(data[i][17]); 
      if (!isNaN(rating)) {
        totalRating += rating;
        count++;
      }
    }

    if (count === 0) {
      
      adminIds.forEach(id => sendTelegramMessage(id, null, '📊 Belum ada data rating yang masuk.'));
      return;
    }

    const avg = (totalRating / count).toFixed(2);
    const pesan = `📊 *Laporan Rekap Rating Chatbot*\n\n` +
                  `Jumlah pengaduan dengan rating: ${count}\n` +
                  `Rata-rata rating: ⭐ *${avg} / 5*\n\n` +
                  `Terima kasih atas peningkatan layanan yang terus dijaga 🙌`;


    adminIds.forEach(id => sendTelegramMessage(id, null, pesan));
    log(`📊 Rekap rating dikirim ke admin (${avg} dari ${count} rating)`);
  } catch (err) {
    log(`❌ Error rekapRating: ${err}`);
  }

}

function generateDashboardPremium() {

  const ss = SpreadsheetApp.getActive();
  const sumber = ss.getSheetByName("Data Pengaduan");

  // === Tema Material 3 ===
  const BG = "#f5f7fa";
  const CARD = "#ffffff";
  const BORDER = "#d6d9de";
  const TEXT = "#1a1a1a";
  const HEADER = "#e8eef7";

  // Utility: load/create sheet
  function getOrCreateSheet(name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.clear();
    sh.getRange("A1:Z999").setBackground(BG);
    return sh;
  }

  // Load data
  const data = sumber.getDataRange().getValues();
  if (data.length < 2) return;

  const header = data[0];
  const rows = data.slice(1);

  const colTanggal = header.indexOf("Tanggal");
  const colKategori = header.indexOf("Kategori");
  const colStatus = header.indexOf("Status");
  

  // Base data structures
  const daily = {};
  const weekly = {};
  const monthly = {};
  const perKategori = {};
  const perStatus = {};

  // NEW: weekly per category, monthly per category
  const weeklyKategori = {};
  const monthlyKategori = {};

  rows.forEach(r => {
    if (!r[colTanggal]) return;

    const tgl = new Date(r[colTanggal]);
    const kategori = r[colKategori] || "Lainnya";
    const status = r[colStatus] || "Tidak Ada";
    

    const tglStr = Utilities.formatDate(tgl, "GMT+7", "yyyy-MM-dd");
    const mingguStr = Utilities.formatDate(tgl, "GMT+7", "YYYY-'W'ww");
    const bulanStr = Utilities.formatDate(tgl, "GMT+7", "yyyy-MM");

    // Harian per kategori
    if (!daily[tglStr]) daily[tglStr] = {};
    daily[tglStr][kategori] = (daily[tglStr][kategori] || 0) + 1;

    // Mingguan total
    weekly[mingguStr] = (weekly[mingguStr] || 0) + 1;

    // NEW: Mingguan per kategori
    if (!weeklyKategori[mingguStr]) weeklyKategori[mingguStr] = {};
    weeklyKategori[mingguStr][kategori] =
      (weeklyKategori[mingguStr][kategori] || 0) + 1;

    // Bulanan total
    monthly[bulanStr] = (monthly[bulanStr] || 0) + 1;

    // NEW: Bulanan per kategori
    if (!monthlyKategori[bulanStr]) monthlyKategori[bulanStr] = {};
    monthlyKategori[bulanStr][kategori] =
      (monthlyKategori[bulanStr][kategori] || 0) + 1;

    perKategori[kategori] = (perKategori[kategori] || 0) + 1;
    perStatus[status] = (perStatus[status] || 0) + 1;
  });

  // Utility menulis tabel
  function writeTable(sheet, row, col, data, title = "") {
    const range = sheet.getRange(row, col, data.length, data[0].length);

    range
      .setValues(data)
      .setBackground(CARD)
      .setFontColor(TEXT)
      .setBorder(true, true, true, true, true, true);

    sheet.getRange(row, col, 1, data[0].length)
      .setBackground(HEADER)
      .setFontWeight("bold");

    if (title) {
      sheet.getRange(row - 1, col)
        .setValue(title)
        .setFontWeight("bold")
        .setFontColor(TEXT);
    }

    return range;
  }

  // =====================================================
  // 1. LAPORAN HARIAN (per kategori)
  // =====================================================
  const shDaily = getOrCreateSheet("lap_harian");

  const kategoriList = Array.from(
    new Set(Object.values(daily).flatMap(o => Object.keys(o)))
  );

  const t1 = [["Tanggal", ...kategoriList]];
  Object.keys(daily).sort().forEach(t => {
    t1.push([t, ...kategoriList.map(k => daily[t][k] || 0)]);
  });

  const rDaily = writeTable(shDaily, 3, 1, t1, "Pengaduan Harian per Kategori");

  shDaily.insertChart(
    shDaily.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(rDaily)
      .setOption("isStacked", true)
      .setPosition(rDaily.getRow() + rDaily.getNumRows() + 2, 1, 0, 0)
      .build()
  );

  // =====================================================
  // 2. LAPORAN MINGGUAN (total)
  // =====================================================
  const shWeekly = getOrCreateSheet("lap_mingguan");

  const t2 = [["Minggu", "Jumlah"]];
  Object.keys(weekly).sort().forEach(m => t2.push([m, weekly[m]]));

  const rWeekly = writeTable(shWeekly, 3, 1, t2, "Rekap Mingguan");

  shWeekly.insertChart(
    shWeekly.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(rWeekly)
      .setPosition(rWeekly.getRow() + rWeekly.getNumRows() + 2, 1, 0, 0)
      .build()
  );

  // =====================================================
  // 3. LAPORAN BULANAN (total)
  // =====================================================
  const shMonthly = getOrCreateSheet("lap_bulanan");

  const t3 = [["Bulan", "Jumlah"]];
  Object.keys(monthly).sort().forEach(b => t3.push([b, monthly[b]]));

  const rMonthly = writeTable(shMonthly, 3, 1, t3, "Rekap Bulanan");

  shMonthly.insertChart(
    shMonthly.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(rMonthly)
      .setPosition(rMonthly.getRow() + rMonthly.getNumRows() + 2, 1, 0, 0)
      .build()
  );

  // =====================================================
  // 4. LAPORAN KATEGORI (total)
  // =====================================================
  const shKategori = getOrCreateSheet("lap_kategori");

  const t4 = [["Kategori", "Jumlah"]];
  Object.keys(perKategori).forEach(k => t4.push([k, perKategori[k]]));

  const rKategori = writeTable(shKategori, 3, 1, t4, "Distribusi Kategori");

  shKategori.insertChart(
    shKategori.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(rKategori)
      .setOption("pieHole", 0.35)
      .setPosition(rKategori.getRow() + rKategori.getNumRows() + 2, 1, 0, 0)
      .build()
  );

  // =====================================================
  // 5. LAPORAN STATUS
  // =====================================================
  const shStatus = getOrCreateSheet("lap_status");

  const t5 = [["Status", "Jumlah"]];
  Object.keys(perStatus).forEach(s => t5.push([s, perStatus[s]]));

  const rStatus = writeTable(shStatus, 3, 1, t5, "Status Pengaduan");

  shStatus.insertChart(
    shStatus.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(rStatus)
      .setPosition(rStatus.getRow() + rStatus.getNumRows() + 2, 1, 0, 0)
      .build()
  );

  // =====================================================
  // 6. NEW — LAPORAN MINGGUAN PER KATEGORI
  // =====================================================
  const shWeeklyCat = getOrCreateSheet("lap_mingguan_kategori");

  const kategoriAll = Array.from(
    new Set(rows.map(r => r[colKategori] || "Lainnya"))
  );

  const t6 = [["Minggu", ...kategoriAll]];
  Object.keys(weeklyKategori).sort().forEach(m => {
    t6.push([
      m,
      ...kategoriAll.map(k => weeklyKategori[m][k] || 0)
    ]);
  });

  const rWeeklyCat = writeTable(shWeeklyCat, 3, 1, t6, "Rekap Mingguan per Kategori");

  shWeeklyCat.insertChart(
    shWeeklyCat.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(rWeeklyCat)
      .setOption("isStacked", true)
      .setPosition(rWeeklyCat.getRow() + rWeeklyCat.getNumRows() + 2, 1, 0, 0)
      .build()
  );

  // =====================================================
  // 7. NEW — LAPORAN BULANAN PER KATEGORI
  // =====================================================
  const shMonthlyCat = getOrCreateSheet("lap_bulanan_kategori");

  const t7 = [["Bulan", ...kategoriAll]];
  Object.keys(monthlyKategori).sort().forEach(b => {
    t7.push([
      b,
      ...kategoriAll.map(k => monthlyKategori[b][k] || 0)
    ]);
  });

  const rMonthlyCat = writeTable(shMonthlyCat, 3, 1, t7, "Rekap Bulanan per Kategori");

  shMonthlyCat.insertChart(
    shMonthlyCat.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(rMonthlyCat)
      .setOption("isStacked", true)
      .setPosition(rMonthlyCat.getRow() + rMonthlyCat.getNumRows() + 2, 1, 0, 0)
      .build()
  );

}

// =====================================================
// MENU & TRIGGER
// =====================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Laporan")
    .addItem("Perbarui Laporan Sekarang", "generateDashboardPremium")
    .addSeparator()
    .addItem("Aktifkan Update Otomatis (1 Jam)", "enableAutoUpdate")
    .addItem("Nonaktifkan Update Otomatis", "disableAutoUpdate")
    .addToUi();
}

function enableAutoUpdate() {
  disableAutoUpdate();

  ScriptApp.newTrigger("generateDashboardPremium")
    .timeBased()
    .everyHours(1)
    .create();

  SpreadsheetApp.getUi().alert("Update otomatis setiap 1 jam telah diaktifkan.");
}

function disableAutoUpdate() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "generateDashboardPremium") {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function checkTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log("Trigger aktif: " + triggers.map(t => t.getHandlerFunction()));
}
