const TELEBOT_TOKEN = 'xxxx:xxxx';
// const TELEGROUP_ID = xxxx;
const TELEGROUP_ID = "xxx";
const SPREADSHEET_ID = "xxxx";
const NAMA_TAB_UTAMA = "xxx";
const WA_PHONE_NUMBER_ID = "xxxx";
const WA_ACCESS_TOKEN = "xxxx";
const WA_TEMPLATE_NAME = "notifikasi_tiket";
const WA_WEBHOOK_VERIFY_TOKEN = "xxxx";
const SKM_URL = "xx";

function default_keyboard(row) {
    return {
        inline_keyboard: [
            [
                { text: "Sekretariat", callback_data: `PROSES_SEK_${row}` },
                { text: "Ekonomi Desa", callback_data: `PROSES_EKO_${row}` }
            ],
            [
                { text: "Pemberdayaan", callback_data: `PROSES_PMD_${row}` },
                { text: "Keuangan & Aset", callback_data: `PROSES_KEU_${row}` }
            ],
            [
                { text: "Pemerintahan Desa", callback_data: `PROSES_PEM_${row}` }
            ],
            [
                { text: "❌ Tolak Laporan", callback_data: `TANYA_TOLAK_${row}` }
            ]
        ]
    };
}

function generateTicketId(counter) {
    const date = new Date();
    const year = date.getFullYear();
    const month = ("0" + (date.getMonth() + 1)).slice(-2);
    const day = ("0" + date.getDate()).slice(-2);

    // Random pendek (3–4 karakter)
    const randomPart = Math.random().toString(36).substring(2, 6).toUpperCase();

    // ID gabungan
    const idTicket = `PMD-${year}${month}${day}-${counter}-${randomPart}`;
    return idTicket;
}

// function onFormSubmit(e) {
//   try {
//     const sheet = e.range.getSheet();
//     const row = e.range.getRow();
//     const idTicket = generateTicketId(row);

//     const kolomID = 10; 
//     const kolomStatus = 11; 
//     sheet.getRange(row, kolomID).setValue(idTicket);
//     sheet.getRange(row, kolomStatus).setValue("Menunggu Tindakan");

//     const nama            = e.namedValues['Nama Lengkap'] ? e.namedValues['Nama Lengkap'][0] : 'KOSONG';
//     const nik             = e.namedValues['NIK (Nomor Induk Kependudukan)'] ? e.namedValues['NIK (Nomor Induk Kependudukan)'][0] : 'KOSONG';
//     const nama_instansi   = e.namedValues['Nama Instansi'] ? e.namedValues['Nama Instansi'][0] : 'KOSONG';
//     const jenis_keperluan = e.namedValues['Keperluan Kunjungan'] ? e.namedValues['Keperluan Kunjungan'][0] : 'KOSONG';
//     const uraian          = e.namedValues['Uraian Singkat atau Detail Pengaduan'] ? e.namedValues['Uraian Singkat atau Detail Pengaduan'][0] : 'KOSONG';

//     const noWaRaw         = e.namedValues['Nomor WhatsApp atau Telepon'] ? e.namedValues['Nomor WhatsApp atau Telepon'][0] : '';

//     sendTelegramAdmin(idTicket, nama, nik, nama_instansi, jenis_keperluan, uraian, row);

//     if (noWaRaw && noWaRaw !== 'KOSONG') {
//       const cleanNoWa = formatNomorWA(noWaRaw);
//       sendWhatsAppUser(cleanNoWa, nama, idTicket);
//     } else {
//       Logger.log(`Baris ${row}: Nomor WA tidak ditemukan/kosong. Lewati notifikasi WA.`);
//     }

//   } catch (error) {
//     Logger.log("Error fatal pada fungsi onFormSubmit: " + error.toString());
//   }
// }

function onFormSubmit(e) {
    try {
        const sheetProd = e.range.getSheet();
        const rowProd = e.range.getRow();

        const nama = e.namedValues['Nama Lengkap'] ? e.namedValues['Nama Lengkap'][0] : 'KOSONG';

        if (nama.includes("[TEST]") || nama.includes("DEV-")) {


            const DEV_SPREADSHEET_ID = "1w3VyX6waW6cl1rBy9XYeoX1ID3rnCH6nNLLKm_w-RU4";
            const ssDev = SpreadsheetApp.openById(DEV_SPREADSHEET_ID);
            const sheetDev = ssDev.getSheetByName(sheetProd.getName());

            // Pindahkan baris data ke sheet DEV
            const rowData = e.values;
            sheetDev.appendRow(rowData);
            const rowDev = sheetDev.getLastRow();

            sheetProd.deleteRow(rowProd);
            Utilities.sleep(1000);

            try {
                const formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
                if (formUrl) {
                    const form = FormApp.openByUrl(formUrl);
                    const formResponses = form.getResponses();

                    if (formResponses.length > 0) {
                        // Mengambil respon paling terakhir yang masuk
                        const lastResponse = formResponses[formResponses.length - 1];

                        // Validasi tambahan: Pastikan respon yang dihapus memang benar data TEST
                        const itemResponses = lastResponse.getItemResponses();
                        let isTestData = false;

                        for (let i = 0; i < itemResponses.length; i++) {
                            if (itemResponses[i].getItem().getTitle() === 'Nama Lengkap') {
                                const val = itemResponses[i].getResponse();
                                if (val.includes("[TEST]") || val.includes("DEV-")) {
                                    isTestData = true;
                                }
                                break;
                            }
                        }

                        // Jika tervalidasi benar data test, hapus dari database internal Form
                        if (isTestData) {
                            form.deleteResponse(lastResponse.getId());
                            Logger.log("[DEV] Sukses menghapus data test dari internal Google Form.");
                        }
                    }
                }
            } catch (err) {
                Logger.log("Gagal membersihkan database internal Form: " + err.message);
            }

            const idTicketDev = generateTicketId(rowDev); // Menggunakan baris di sheet dev
            const kolomID = 10;
            const kolomStatus = 11;
            sheetDev.getRange(rowDev, kolomID).setValue(idTicketDev);
            sheetDev.getRange(rowDev, kolomStatus).setValue("Menunggu Tindakan");

            const nik = e.namedValues['NIK (Nomor Induk Kependudukan)'] ? e.namedValues['NIK (Nomor Induk Kependudukan)'][0] : 'KOSONG';
            const nama_instansi = e.namedValues['Nama Instansi'] ? e.namedValues['Nama Instansi'][0] : 'KOSONG';
            const jenis_keperluan = e.namedValues['Keperluan Kunjungan'] ? e.namedValues['Keperluan Kunjungan'][0] : 'KOSONG';
            const uraian = e.namedValues['Uraian Singkat atau Detail Pengaduan'] ? e.namedValues['Uraian Singkat atau Detail Pengaduan'][0] : 'KOSONG';
            const noWaRaw = e.namedValues['Nomor WhatsApp atau Telepon'] ? e.namedValues['Nomor WhatsApp atau Telepon'][0] : '';

            sendTelegramAdmin(idTicketDev, nama, nik, nama_instansi, jenis_keperluan, uraian, rowDev);

            if (noWaRaw && noWaRaw !== 'KOSONG') {
                const cleanNoWa = formatNomorWA(noWaRaw);
                sendWhatsAppUser(cleanNoWa, nama, idTicketDev); // Kirim ke nomor testing Anda
            }

            Logger.log(`[DEV] Alur pengujian berhasil dieksekusi untuk baris ${rowDev}`);

        } else {


            const idTicket = generateTicketId(rowProd);
            const kolomID = 10;
            const kolomStatus = 11;
            sheetProd.getRange(rowProd, kolomID).setValue(idTicket);
            sheetProd.getRange(rowProd, kolomStatus).setValue("Menunggu Tindakan");

            const nik = e.namedValues['NIK (Nomor Induk Kependudukan)'] ? e.namedValues['NIK (Nomor Induk Kependudukan)'][0] : 'KOSONG';
            const nama_instansi = e.namedValues['Nama Instansi'] ? e.namedValues['Nama Instansi'][0] : 'KOSONG';
            const jenis_keperluan = e.namedValues['Keperluan Kunjungan'] ? e.namedValues['Keperluan Kunjungan'][0] : 'KOSONG';
            const uraian = e.namedValues['Uraian Singkat atau Detail Pengaduan'] ? e.namedValues['Uraian Singkat atau Detail Pengaduan'][0] : 'KOSONG';
            const noWaRaw = e.namedValues['Nomor WhatsApp atau Telepon'] ? e.namedValues['Nomor WhatsApp atau Telepon'][0] : '';

            sendTelegramAdmin(idTicket, nama, nik, nama_instansi, jenis_keperluan, uraian, rowProd);

            if (noWaRaw && noWaRaw !== 'KOSONG') {
                const cleanNoWa = formatNomorWA(noWaRaw);
                sendWhatsAppUser(cleanNoWa, nama, idTicket);
            }
        }

    } catch (error) {
        Logger.log("Error fatal pada fungsi onFormSubmit: " + error.toString());
    }
}

function sendTelegramAdmin(idTicket, nama, nik, nama_instansi, jenis_keperluan, uraian, row) {
    const pesanText =
        `🚨 *Laporan/Tamu Baru Masuk!* 🚨

  📌 *ID Tiket:* ${idTicket}
  👤 *Nama:* ${nama}
  🆔 *NIK:* ${nik}
  🏢 *Instansi:* ${nama_instansi}
  🎯 *Keperluan:* ${jenis_keperluan}
  📝 *Uraian:* ${uraian}

  Silakan cek detail lebih lanjut di *Dashboard* atau *Google Sheets*.`;

    const url = `https://api.telegram.org/bot${TELEBOT_TOKEN}/sendMessage`;
    const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
            chat_id: TELEGROUP_ID,
            text: pesanText,
            parse_mode: 'Markdown',
            reply_markup: default_keyboard(row)
        }),
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    Logger.log("Log Telegram: " + response.getContentText());
}

function sendWhatsAppUser(noWa, name, id_ticket) {
    const url = `https://graph.facebook.com/v25.0/${WA_PHONE_NUMBER_ID}/messages`;

    const payload = {
        "messaging_product": "whatsapp",
        "to": noWa,
        "type": "template",
        "template": {
            "name": WA_TEMPLATE_NAME,
            "language": { "code": "id" },
            "components": [
                {
                    "type": "body",
                    "parameters": [
                        { "type": "text", "text": name },
                        { "type": "text", "text": id_ticket }
                    ]
                }
            ]
        }
    };

    //  const payload = {
    //   "messaging_product": "whatsapp",
    //   "to": noWa,
    //   "type": "template",
    //   "template": {
    //     "name": WA_TEMPLATE_NAME,
    //     "language": { "code": "en_US" }
    //   }
    // };

    const options = {
        "method": "post",
        "contentType": "application/json",
        "headers": {
            "Authorization": "Bearer " + WA_ACCESS_TOKEN
        },
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch(url, options);
    Logger.log("Log WhatsApp: " + response.getContentText());
}

function formatNomorWA(phone) {
    let cleaned = phone.toString().replace(/\D/g, '');
    if (cleaned.startsWith('0')) {
        cleaned = '62' + cleaned.slice(1);
    } else if (cleaned.startsWith('8')) {
        cleaned = '62' + cleaned;
    }
    return cleaned;
}


// function onFormSubmit(e) {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

//   const row = e.range.getRow();

//   const date = new Date();
//   const year = date.getFullYear();
//   const month = ("0" + (date.getMonth() + 1)).slice(-2);
//   const day = ("0" + date.getDate()).slice(-2);
//   const idTicket = `PMD-${year}${month}${day}-${row}`;

//   const kolomID = 10; 
//   const kolomStatus = 11; 

//   sheet.getRange(row, kolomID).setValue(idTicket);
//   sheet.getRange(row, kolomStatus).setValue("Menunggu Tindakan");

//   const nama = e.namedValues['Nama Lengkap'] ? e.namedValues['Nama Lengkap'][0] : 'KOSONG';
//   const nik = e.namedValues['NIK (Nomor Induk Kependudukan)'] ? e.namedValues['NIK (Nomor Induk Kependudukan)'][0] : 'KOSONG';
//   const nama_instansi = e.namedValues['Nama Instansi'] ? e.namedValues['Nama Instansi'][0] : 'KOSONG';
//   const jenis_keperluan = e.namedValues['Keperluan Kunjungan'] ? e.namedValues['Keperluan Kunjungan'][0] : 'KOSONG';
//   const uraian = e.namedValues['Uraian Singkat atau Detail Pengaduan'] ? e.namedValues['Uraian Singkat atau Detail Pengaduan'][0] : 'KOSONG';

//   const pesanText = 
//   `🚨 *Laporan/Tamu Baru Masuk!* 🚨

//   📌 *ID Tiket:* ${idTicket}
//   👤 *Nama:* ${nama}
//   🆔 *NIK:* ${nik}
//   🏢 *Instansi:* ${nama_instansi}
//   🎯 *Keperluan:* ${jenis_keperluan}
//   📝 *Uraian:* ${uraian}

//   Silakan cek detail lebih lanjut di *Dashboard* atau *Google Sheets*.`;

//   const url = `https://api.telegram.org/bot${TELEBOT_TOKEN}/sendMessage`;
//   const options = {
//     method: 'post',
//     contentType: 'application/json',
//     payload: JSON.stringify({
//       chat_id: TELEGROUP_ID,
//       text: pesanText,
//       parse_mode: 'Markdown',
//       reply_markup: default_keyboard(row) 
//     })
//   };

//   const response = UrlFetchApp.fetch(url, options);
//   Logger.log(response.getContentText());
// }

function activateTelegramWebhook() {
    const token = TELEBOT_TOKEN;
    const urlWebApp = 'https://script.google.com/macros/s/AKfycbyWzBMUlKNwPyil99TQh1B9UlbwGa3OHMdvKqdoHcsYYpc5Knxz5n3viDGwVj-dvlXP/exec';

    const url = `https://api.telegram.org/bot${token}/setWebhook?url=${urlWebApp}&drop_pending_updates=true`;
    const response = UrlFetchApp.fetch(url);
    Logger.log(response.getContentText());
}

function doPost(e) {
    try {
        if (!e?.postData?.contents) {
            return ContentService.createTextOutput("OK");
        }

        const rawData = e.postData.contents;
        const data = JSON.parse(rawData);

        if (data.update_id) {
            return handleTelegramWebhook(data, rawData);
        }

        if (data.object === "whatsapp_business_account" || data.entry) {
            return handleWhatsAppWebhook(data, rawData);
        }


        return ContentService.createTextOutput("OK");

    } catch (err) {
        console.error("Gateway Router Error:", err);
        return ContentService.createTextOutput("OK");
    }
}

function handleTelegramWebhook(data, rawData) {
    const updateId = data.update_id;
    const cache = CacheService.getScriptCache();
    const cacheKey = "processed_" + updateId;

    let isDuplicate = false;
    const lock = LockService.getScriptLock();
    try {
        if (lock.tryLock(3000)) {
            if (cache.get(cacheKey)) {
                isDuplicate = true;
            } else {
                cache.put(cacheKey, "true", 7200);
            }
        }
    } catch (e) {
        console.error("Lock/Cache Error in handleTelegramWebhook (proceeding anyway):", e);
    } finally {
        try {
            if (lock.hasLock()) {
                lock.releaseLock();
            }
        } catch (el) {
            console.error("Error releasing lock:", el);
        }
    }

    if (isDuplicate) {
        return ContentService.createTextOutput("OK");
    }

    if (data.callback_query) {
        UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEBOT_TOKEN}/answerCallbackQuery`, {
            method: 'post',
            contentType: 'application/json',
            muteHttpExceptions: true,
            payload: JSON.stringify({ callback_query_id: data.callback_query.id, text: "Memproses..." })
        });
    }

    prosesPendingUpdate(updateId, data);

    return ContentService.createTextOutput("OK");
}

function handleWhatsAppWebhook(data, rawData) {
    let waUniqueId = "";

    try {
        const changeValue = data.entry?.[0]?.changes?.[0]?.value;

        if (changeValue?.messages?.[0]) {
            waUniqueId = changeValue.messages[0].id;
        } else if (changeValue?.statuses?.[0]) {
            waUniqueId = changeValue.statuses[0].id;
        } else {
            waUniqueId = "wa_init_" + new Date().getTime();
        }
    } catch (err) {
        waUniqueId = "wa_err_" + new Date().getTime();
    }

    const cache = CacheService.getScriptCache();
    const cacheKey = "processed_wa_" + waUniqueId;

    let isDuplicate = false;
    const lock = LockService.getScriptLock();
    try {
        if (lock.tryLock(3000)) {
            if (cache.get(cacheKey)) {
                isDuplicate = true;
            } else {
                cache.put(cacheKey, "true", 7200);
            }
        }
    } catch (e) {
        console.error("Lock/Cache Error in handleWhatsAppWebhook (proceeding anyway):", e);
    } finally {
        try {
            if (lock.hasLock()) {
                lock.releaseLock();
            }
        } catch (el) {
            console.error("Error releasing lock:", el);
        }
    }

    if (isDuplicate) {
        return ContentService.createTextOutput("OK");
    }

    prosesPendingWhatsAppUpdate(waUniqueId, data);

    return ContentService.createTextOutput("OK");
}

// function doPost(e) {
//   try {
//     if (!e?.postData?.contents) {
//       return ContentService.createTextOutput("OK");
//     }

//     const rawData = e.postData.contents;
//     const data = JSON.parse(rawData);
//     const updateId = data.update_id;

//     const cache = PropertiesService.getScriptProperties();

//     if (cache.getProperty("processed_" + updateId)) {
//       return ContentService.createTextOutput("OK");
//     }

//     cache.setProperty("processed_" + updateId, "true");

//     if (data.callback_query) {
//       UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEBOT_TOKEN}/answerCallbackQuery`, {
//         method: 'post',
//         contentType: 'application/json',
//         muteHttpExceptions: true,
//         payload: JSON.stringify({ callback_query_id: data.callback_query.id, text: "Memproses..." })
//       });
//     }

//     // Simpan untuk diproses
//     cache.setProperty("pending_" + updateId, rawData);

//     // Proses di fungsi terpisah
//     prosesPendingUpdate(updateId);

//     return ContentService.createTextOutput("OK");

//   } catch (err) {
//     console.error("doPost Error:", err);
//     return ContentService.createTextOutput("OK");
//   }
// }

function prosesPencarianTiket(userPhone, ticketId) {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const data = sheet.getDataRange().getValues();

    const colIndexTiket = 9
    const colIndexStatus = 10
    const colIndexNama = 1
    const colIndexUlasan = 12
    const colIndexBidangPJ = 11
    const colIndexKeperluan = 7

    let rowFound = -1;

    for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][colIndexTiket].toString().toUpperCase() === ticketId.toUpperCase()) {
            rowFound = i;
            break;
        }
    }

    if (rowFound !== -1) {
        const namaUser = data[rowFound][colIndexNama] || "Pemohon";
        const statusSaatIni = data[rowFound][colIndexStatus] || "Menunggu Tindakan";
        const ulasanAdmin = data[rowFound][colIndexUlasan] || "Belum ada ulasan dari petugas.";
        const bidangPJ = data[rowFound][colIndexBidangPJ] || "Belum ditentukan";
        const keperluan = data[rowFound][colIndexKeperluan] || "Tidak ada keterangan";

        let pesanBalasan = `📄 *DETAIL STATUS LAPORAN* 📄\n\n` +
            `🎫 *ID Tiket:* ${ticketId}\n` +
            `👤 *Nama:* ${namaUser}\n` +
            `🏢 *Bidang Penanggung Jawab:* ${bidangPJ}\n` +
            `📌 *Keperluan Kunjungan:* ${keperluan}\n` +
            `📊 *Status Laporan:* _${statusSaatIni}_\n` +
            `📝 *Keterangan/Ulasan:* ${ulasanAdmin}\n\n`;

        if (statusSaatIni.toUpperCase() === "DITOLAK") {
            pesanBalasan += `🙏 Kami mohon maaf, laporan Anda tidak dapat diproses lebih lanjut.\n\n` +
                `💡 Untuk membantu kami meningkatkan pelayanan, mohon kesediaan Anda mengisi *Survei Kepuasan Masyarakat (SKM)* melalui tautan berikut:\n` +
                `${SKM_URL}\n\n` +
                `Jika masih ada hal yang perlu ditanyakan, silakan menghubungi petugas kami agar mendapatkan penjelasan lebih lanjut. Terima kasih atas pengertian dan kerja samanya.`;
        }
        kirimWaTeksBiasa(userPhone, pesanBalasan);
    } else {
        kirimWaTeksBiasa(userPhone, `❌ *Tiket Tidak Ditemukan*\n\nMaaf, nomor tiket *${ticketId}* tidak terdaftar di sistem kami. Mohon periksa kembali karakter dan ejaan tiket Anda.`);
    }
}

function kirimWaTeksBiasa(noWa, teks) {
    const url = `https://graph.facebook.com/v25.0/${WA_PHONE_NUMBER_ID}/messages`;
    const payload = {
        "messaging_product": "whatsapp",
        "recipient_type": "individual",
        "to": noWa,
        "type": "text",
        "text": {
            "preview_url": true,
            "body": teks
        }
    };

    const options = {
        "method": "post",
        "contentType": "application/json",
        "headers": {
            "Authorization": "Bearer " + WA_ACCESS_TOKEN
        },
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
    };
    UrlFetchApp.fetch(url, options);
}

function prosesPendingWhatsAppUpdate(waUniqueId, data) {
    if (!data) return;

    try {
        const changeValue = data.entry[0].changes[0].value;

        if (changeValue.messages && changeValue.messages[0]) {
            const messageObject = changeValue.messages[0];
            const userPhone = messageObject.from;

            if (messageObject.type === "text") {
                const userText = messageObject.text.body.trim();

                if (userText.toUpperCase().startsWith("CEK")) {
                    const ticketId = userText.replace(/CEK/i, "").trim();
                    if (ticketId) {
                        prosesPencarianTiket(userPhone, ticketId);
                    } else {
                        kirimWaTeksBiasa(userPhone, "⚠️ Format salah.\nSilakan ketik: *CEK [Nomor Tiket]*");
                    }
                } else {
                    const menuSapaan = `👋 Halo! Selamat datang di Layanan Si Buku Tamu Online DPMD Kabupaten Muara Enim.\n\nUntuk mengecek status perkembangan laporan Anda, silakan ketik:\n*CEK [Nomor Tiket]*`;
                    kirimWaTeksBiasa(userPhone, menuSapaan);
                }
            }
        }

    } catch (err) {
        console.error("Gagal memproses Pending WhatsApp Update ID " + waUniqueId + ": ", err);
    }
}

function prosesPendingUpdate(updateId, data) {
    let logSheet = null;
    const cache = CacheService.getScriptCache();
    const startTime = new Date();

    let typeLog = "UNKNOWN";
    let aksiLog = "UNKNOWN";
    let rowLog = "";
    let statusLog = "START";
    let detailLog = "Proses dimulai";

    try {
        if (!data) return;
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

        logSheet = ss.getSheetByName("Debug_Log") || ss.insertSheet("Debug_Log");
        if (logSheet.getLastRow() === 0) {
            logSheet.appendRow(["Waktu", "Tipe", "UpdateID", "Aksi", "Row", "Status", "Detail"]);
        }

        const sheet = ss.getSheetByName(NAMA_TAB_UTAMA);
        if (!sheet) throw new Error("Sheet utama tidak ditemukan");

        const isCallback = !!data.callback_query;
        typeLog = isCallback ? "CALLBACK" : "REPLY";
        const isReply = !isCallback && data.message && data.message.reply_to_message;

        const mapBidang = {
            "SEK": "Bidang Sekretariat",
            "EKO": "Bidang Pengembangan Kerjasama Ekonomi Desa",
            "PMD": "Bidang Pemberdayaan Masyarakat dan Kelembagaan",
            "KEU": "Bidang Keuangan dan Aset Desa",
            "PEM": "Bidang Pemerintahan dan Penataan Wilayah Desa"
        };

        // ==================== CALLBACK QUERY ====================
        if (isCallback) {
            const cq = data.callback_query;
            const callbackData = cq.data;
            const parts = callbackData.split('_');
            aksiLog = parts[0];
            rowLog = parseInt(parts[parts.length - 1]);

            console.log(`[CALLBACK] Processing callbackData: ${callbackData}`);

            if (isNaN(rowLog) || rowLog < 2) throw new Error("Row tidak valid");

            // ==================== BATAL ====================
            if (aksiLog === "BATAL") {
                sheet.getRange(rowLog, 11).setValue("Menunggu Tindakan");
                sheet.getRange(rowLog, 12).setValue("");
                UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEBOT_TOKEN}/editMessageReplyMarkup`, {
                    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
                    payload: JSON.stringify({
                        chat_id: cq.message.chat.id,
                        message_id: cq.message.message_id,
                        reply_markup: default_keyboard(rowLog)
                    })
                });
                statusLog = "SUCCESS";
                detailLog = "Tombol dibatalkan";
                return;
            }

            // ==================== PROSES (Pilih Bidang) ====================
            if (aksiLog === "PROSES") {
                const bidangCode = parts[1];
                const namaBidang = mapBidang[bidangCode] || bidangCode;

                sheet.getRange(rowLog, 11).setValue("Diproses");
                sheet.getRange(rowLog, 12).setValue(namaBidang);

                const pesanUpdate = `${cq.message.text}\n\n✅ *STATUS: Diproses*\nOleh: ${namaBidang}\nPetugas: ${cq.from.first_name}`;

                const keyboard = {
                    inline_keyboard: [
                        [{ text: "✅ Tandai Selesai", callback_data: `TANYA_SELESAI_${rowLog}` }],
                        [{ text: "❌ Batal", callback_data: `BATAL_${rowLog}` }]
                    ]
                };

                UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEBOT_TOKEN}/editMessageText`, {
                    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
                    payload: JSON.stringify({
                        chat_id: cq.message.chat.id,
                        message_id: cq.message.message_id,
                        text: pesanUpdate,
                        parse_mode: 'Markdown',
                        reply_markup: keyboard
                    })
                });

                statusLog = "SUCCESS";
                detailLog = `Diproses oleh ${namaBidang}`;
            }

            // ==================== TANYA (Tolak / Selesai) ====================
            else if (aksiLog === "TANYA") {
                const jenis = parts[1];
                const rowStr = parts[2];
                const idTicket = sheet.getRange(parseInt(rowStr), 10).getValue() || "Unknown";
                const kataTugas = jenis === "TOLAK" ? "alasan penolakan" : "catatan penyelesaian";
                const username = cq.from.username ? "@" + cq.from.username : cq.from.first_name;
                const originalMessageId = cq.message.message_id;

                const pesanUpdate = `${cq.message.text}\n\n✅ *STATUS: Ditolak*\nOleh: ${username}\n`;

                const keyboard = {
                    inline_keyboard: [
                        [{ text: "❌ Batal", callback_data: `BATAL_${rowLog}` }]
                    ]
                };

                UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEBOT_TOKEN}/editMessageText`, {
                    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
                    payload: JSON.stringify({
                        chat_id: cq.message.chat.id,
                        message_id: cq.message.message_id,
                        text: pesanUpdate,
                        parse_mode: 'Markdown',
                        reply_markup: keyboard
                    })
                });

                const replyMsg = `[#ROW_${rowStr}_${jenis}] [#MAIN_${originalMessageId}] ${username}, silakan balas (reply) pesan ini dengan *${kataTugas}* untuk laporan ini.\nTiket: ${idTicket}`;

                UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEBOT_TOKEN}/sendMessage`, {
                    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
                    payload: JSON.stringify({
                        chat_id: TELEGROUP_ID,
                        text: replyMsg,
                        parse_mode: 'Markdown',
                        reply_to_message_id: originalMessageId,
                        reply_markup: { force_reply: true, selective: true }
                    })
                });

                statusLog = "SUCCESS";
                detailLog = `Instruksi ${jenis} dikirim | MAIN_ID: ${originalMessageId}`;
            }
        }

        // ==================== REPLY MESSAGE ====================
        else if (isReply) {
            const repliedText = data.message.reply_to_message ? data.message.reply_to_message.text || "" : "";
            const matchRow = repliedText.match(/#ROW_(\d+)_([A-Z]+)/);
            const matchMain = repliedText.match(/#MAIN_(\d+)/);

            if (matchRow && matchMain) {
                rowLog = parseInt(matchRow[1]);
                const jenis = matchRow[2];
                aksiLog = jenis;
                const originalMessageId = parseInt(matchMain[1]);

                const teksBalasan = data.message.text;
                const pengisi = data.message.from.first_name;
                const idTicket = sheet.getRange(rowLog, 10).getValue() || "";
                const chatId = data.message.chat.id;

                console.log(`[REPLY] Processing ${jenis} for row ${rowLog}, ticket: ${idTicket}`);

                if (jenis === "TOLAK") {
                    sheet.getRange(rowLog, 11).setValue("Ditolak");
                    sheet.getRange(rowLog, 13).setValue(teksBalasan);

                } else if (jenis === "SELESAI") {
                    sheet.getRange(rowLog, 11).setValue("Selesai");
                    sheet.getRange(rowLog, 13).setValue(teksBalasan);

                    const colIndexTiket = 10;
                    const colIndexWA = 4;

                    const idTiket = sheet.getRange(rowLog, colIndexTiket).getValue();
                    const noWaRaw = sheet.getRange(rowLog, colIndexWA).getValue();

                    if (noWaRaw && noWaRaw !== "KOSONG" && noWaRaw !== "") {
                        const noWaBersih = formatNomorWA(noWaRaw);
                        const pesanSelesai =
                            `Halo Bapak/Ibu,

        Kami informasikan bahwa laporan/pengaduan Anda dengan nomor tiket *${idTiket}* telah *SELESAI* ditindaklanjuti.

        Sebagai upaya peningkatan kualitas pelayanan publik, silakan Bapak/Ibu mengisi Survei Kepuasan Masyarakat (SKM) melalui tautan berikut:
        🔗 ${SKM_URL}

        Setelah mengisi survei tersebut, Anda dapat melihat detail penyelesaian dengan mengetik:
        *CEK ${idTiket}*

        Terima kasih atas partisipasi Anda telah menggunakan SI Buku Tamu Online DPMD Kabupaten Muara Enim. 🙏`;
                        kirimWaTeksBiasa(noWaBersih, pesanSelesai);
                        console.log(`[Telegram Trigger] Notifikasi Selesai & SKM berhasil dikirim ke WA: ${noWaBersih}`);

                    } else {
                        console.warn(`[Telegram Trigger] Tiket ${idTiket} diselesaikan, namun nomor WA pengguna tidak ditemukan.`);
                    }
                }

                try {
                    const hideRes = UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEBOT_TOKEN}/editMessageReplyMarkup`, {
                        method: 'post',
                        contentType: 'application/json',
                        muteHttpExceptions: true,
                        payload: JSON.stringify({
                            chat_id: chatId,
                            message_id: originalMessageId,
                            reply_markup: { inline_keyboard: [] }
                        })
                    });
                    console.log(`Hide markup result: ${hideRes.getResponseCode()}`);
                } catch (err) {
                    console.error("Error hiding markup:", err);
                }

                // Konfirmasi
                UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEBOT_TOKEN}/sendMessage`, {
                    method: 'post',
                    contentType: 'application/json',
                    muteHttpExceptions: true,
                    payload: JSON.stringify({
                        chat_id: chatId,
                        text: `✅ Laporan ${idTicket} berhasil ${jenis === "TOLAK" ? "DITOLAK" : "DISELESAIKAN"} oleh ${pengisi}.`,
                        parse_mode: "Markdown"
                    })
                });

                statusLog = "SUCCESS";
                detailLog = `Laporan berhasil ${jenis === "TOLAK" ? "DITOLAK" : "DISELESAIKAN"} oleh ${pengisi}`;
            } else {
                statusLog = "NO_MATCH";
                detailLog = "Tidak menemukan #ROW_ atau #MAIN_";
            }
        }

    } catch (err) {
        statusLog = "FAILED";
        detailLog = err.toString().substring(0, 200);
        console.error("Proses Pending Error:", err);
    } finally {
        const duration = new Date() - startTime;
        if (logSheet) {
            logSheet.appendRow([new Date(), typeLog, updateId, aksiLog, rowLog, statusLog, `${detailLog} (Duration: ${duration}ms)`]);
        }
    }
}

function doGet(e) {
    const mode = e.parameter["hub.mode"];
    const token = e.parameter["hub.verify_token"];
    const challenge = e.parameter["hub.challenge"];

    if (mode && token) {
        if (mode === "subscribe" && token === WA_WEBHOOK_VERIFY_TOKEN) {
            return ContentService.createTextOutput(challenge);
        }
    }
    return ContentService.createTextOutput("Forbidden");
}

function webhookTelegramCheck() {
    const token = TELEBOT_TOKEN;
    const url = `https://api.telegram.org/bot${token}/getWebhookInfo`;
    const response = UrlFetchApp.fetch(url);
    console.log(response.getContentText());
}

function deactivateTelegramWebhook() {
    const token = TELEBOT_TOKEN;
    const url = `https://api.telegram.org/bot${token}/setWebhook?url=`;
    const response = UrlFetchApp.fetch(url);
    console.log(response.getContentText());
}

function runTestTelgramButtonClick() {
    const mockEvent = {
        postData: {
            contents: JSON.stringify({
                callback_query: {
                    id: "MOCK_CALLBACK_ID_123",
                    data: "PROSES_SEK_2",
                    message: {
                        message_id: 9999,
                        chat: { id: 123456789 },
                        text: "Laporan Asli Masuk dari Google Form"
                    },
                    from: {
                        first_name: "Admin Tester",
                        username: "tester_admin"
                    }
                }
            })
        }
    };

    console.log("--- MEMULAI TES KLIK TOMBOL (BARIS 10) ---");
    doPost(mockEvent);
    console.log("--- TES SELESAI --- Kunjungi Spreadsheet Anda untuk cek Baris 10 Kolom K & L");
}

function runTestTelgramReply() {
    const mockEvent = {
        postData: {
            contents: JSON.stringify({
                message: {
                    message_id: 8888,
                    chat: { id: 123456789 },
                    text: "Dokumen kurang lengkap, silakan upload ulang.", // Isi alasan penolakan
                    from: { first_name: "Petugas Tester" },
                    reply_to_message: {
                        // Teks ini harus mengandung pola regex [#ROW_10_TOLAK] agar dibaca sistem Anda
                        text: "[#ROW_10_TOLAK] @tester_admin, silakan balas pesan ini dengan alasan penolakan."
                    }
                }
            })
        }
    };

    console.log("--- MEMULAI TES BALAS REPLY (BARIS 10) ---");
    doPost(mockEvent);
    console.log("--- TES SELESAI --- Kunjungi Spreadsheet Anda untuk cek Baris 10 Kolom K & M");
}