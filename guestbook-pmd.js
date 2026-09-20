const TELEBOT_TOKEN = 'xxxx:xxxx';
// const TELEGROUP_ID = xxxx;
const TELEGROUP_ID = "xxx";
const SPREADSHEET_ID = "xxxx";
const DEV_SPREADSHEET_ID = "xx-xxxx";
const NAMA_TAB_UTAMA = "xxx";
const WA_PHONE_NUMBER_ID = "xxxx";
const WA_ACCESS_TOKEN = "xxxx";
const WA_TEMPLATE_NAME = "notifikasi_tiket";
const WA_WEBHOOK_VERIFY_TOKEN = "xxxx";
const SKM_URL = "xx";
const WEBAPP_EXEC_URL = "https://script.google.com/macros/s/AKfycbyWzBMUlKNwPyil99TQh1B9UlbwGa3OHMdvKqdoHcsYYpc5Knxz5n3viDGwVj-dvlXP/exec";

function escapeMarkdown(text) {
    if (text === null || text === undefined) return "";
    return text.toString().replace(/([*_`\[])/g, '\\$1');
}

function sanitizeInputText(text, maxLength = 1000) {
    if (text === null || text === undefined) return "";
    let str = text.toString().trim();
    if (str.length > maxLength) {
        str = str.substring(0, maxLength) + "... [Dipotong]";
    }
    return str;
}

function isValidIndonesianPhone(phone) {
    if (!phone) return false;
    const clean = phone.toString().replace(/\D/g, '');
    return /^628\d{7,12}$/.test(clean);
}

function isValidTicketFormat(ticketId) {
    if (!ticketId) return false;
    return /^PMD-\d{8}-\d+-[A-Z0-9]{3,6}$/i.test(ticketId.trim());
}

function validateRowBounds(sheet, row) {
    if (!sheet || !row || isNaN(row)) return false;
    const r = parseInt(row);
    return r >= 2 && r <= sheet.getLastRow();
}

function default_keyboard(row, isDev = false) {
    const prefix = isDev ? "DEV_" : "";
    return {
        inline_keyboard: [
            [
                { text: "Sekretariat", callback_data: `${prefix}PROSES_SEK_${row}` },
                { text: "Ekonomi Desa", callback_data: `${prefix}PROSES_EKO_${row}` }
            ],
            [
                { text: "Pemberdayaan", callback_data: `${prefix}PROSES_PMD_${row}` },
                { text: "Keuangan & Aset", callback_data: `${prefix}PROSES_KEU_${row}` }
            ],
            [
                { text: "Pemerintahan Desa", callback_data: `${prefix}PROSES_PEM_${row}` }
            ],
            [
                { text: "❌ Tolak Laporan", callback_data: `${prefix}TANYA_TOLAK_${row}` }
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

function onFormSubmit(e) {
    try {
        const sheetProd = e.range.getSheet();
        const rowProd = e.range.getRow();

        const nama = e.namedValues['Nama Lengkap'] ? e.namedValues['Nama Lengkap'][0] : 'KOSONG';

        if (nama.includes("[TEST]") || nama.includes("DEV-")) {

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

            // Sanitasi teks input
            const safeNama = sanitizeInputText(nama, 100);
            const safeNik = sanitizeInputText(nik, 20);
            const safeInstansi = sanitizeInputText(nama_instansi, 150);
            const safeKeperluan = sanitizeInputText(jenis_keperluan, 200);
            const safeUraian = sanitizeInputText(uraian, 1000);

            const telegramMsgIdDev = sendTelegramAdmin(idTicketDev, safeNama, safeNik, safeInstansi, safeKeperluan, safeUraian, rowDev, true);
            if (telegramMsgIdDev) {
                sheetDev.getRange(rowDev, 15).setValue(telegramMsgIdDev); // Kolom O (15) = Telegram_Message_ID
            }

            if (noWaRaw && noWaRaw !== 'KOSONG') {
                const cleanNoWa = formatNomorWA(noWaRaw);
                if (isValidIndonesianPhone(cleanNoWa)) {
                    sendWhatsAppUser(cleanNoWa, safeNama, idTicketDev); // Kirim ke nomor testing Anda
                } else {
                    Logger.log(`[DEV] Baris ${rowDev}: Nomor WA (${cleanNoWa}) tidak memenuhi format 628. Notifikasi WA dilewati.`);
                }
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

            // Sanitasi teks input
            const safeNama = sanitizeInputText(nama, 100);
            const safeNik = sanitizeInputText(nik, 20);
            const safeInstansi = sanitizeInputText(nama_instansi, 150);
            const safeKeperluan = sanitizeInputText(jenis_keperluan, 200);
            const safeUraian = sanitizeInputText(uraian, 1000);

            const telegramMsgId = sendTelegramAdmin(idTicket, safeNama, safeNik, safeInstansi, safeKeperluan, safeUraian, rowProd, false);
            if (telegramMsgId) {
                sheetProd.getRange(rowProd, 15).setValue(telegramMsgId); // Kolom O (15) = Telegram_Message_ID
            }

            if (noWaRaw && noWaRaw !== 'KOSONG') {
                const cleanNoWa = formatNomorWA(noWaRaw);
                if (isValidIndonesianPhone(cleanNoWa)) {
                    sendWhatsAppUser(cleanNoWa, safeNama, idTicket);
                } else {
                    Logger.log(`Baris ${rowProd}: Nomor WA (${cleanNoWa}) tidak memenuhi format 628. Notifikasi WA dilewati.`);
                }
            }
        }

    } catch (error) {
        Logger.log("Error fatal pada fungsi onFormSubmit: " + error.toString());
    }
}

function sendTelegramAdmin(idTicket, nama, nik, nama_instansi, jenis_keperluan, uraian, row, isDev = false) {
    const safeTicket = escapeMarkdown(idTicket);
    const safeNama = escapeMarkdown(nama);
    const safeNik = escapeMarkdown(nik);
    const safeInstansi = escapeMarkdown(nama_instansi);
    const safeKeperluan = escapeMarkdown(jenis_keperluan);
    const safeUraian = escapeMarkdown(uraian);

    const pesanText =
        `🚨 *Laporan/Tamu Baru Masuk!* 🚨\n\n` +
        `  📌 *ID Tiket:* ${safeTicket}\n` +
        `  👤 *Nama:* ${safeNama}\n` +
        `  🆔 *NIK:* ${safeNik}\n` +
        `  🏢 *Instansi:* ${safeInstansi}\n` +
        `  🎯 *Keperluan:* ${safeKeperluan}\n` +
        `  📝 *Uraian:* ${safeUraian}\n\n` +
        `  Silakan cek detail lebih lanjut di *Dashboard* atau *Google Sheets*.`;

    const url = `https://api.telegram.org/bot${TELEBOT_TOKEN}/sendMessage`;
    const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
            chat_id: TELEGROUP_ID,
            text: pesanText,
            parse_mode: 'Markdown',
            reply_markup: default_keyboard(row, isDev)
        }),
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const content = response.getContentText();
    Logger.log("Log Telegram: " + content);

    try {
        const resJson = JSON.parse(content);
        if (resJson.ok && resJson.result && resJson.result.message_id) {
            return resJson.result.message_id;
        }
    } catch (err) {
        Logger.log("Gagal mengambil message_id Telegram: " + err.toString());
    }
    return null;
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

function doPost(e) {
    try {
        if (!e?.postData?.contents) {
            return HtmlService.createHtmlOutput("OK");
        }

        const rawData = e.postData.contents;
        const data = JSON.parse(rawData);

        if (data.update_id) {
            return handleTelegramWebhook(data, rawData);
        }

        if (data.object === "whatsapp_business_account" || data.entry) {
            return handleWhatsAppWebhook(data, rawData);
        }


        return HtmlService.createHtmlOutput("OK");

    } catch (err) {
        console.error("Gateway Router Error:", err);
        return HtmlService.createHtmlOutput("OK");
    }
}

function handleTelegramWebhook(data, rawData) {
    const updateId = data.update_id;
    const cache = CacheService.getScriptCache();
    const cacheKey = "tg_" + updateId;

    try {
        if (cache.get(cacheKey)) {
            return HtmlService.createHtmlOutput("OK");
        }
        cache.put(cacheKey, "1", 21600); // 6 jam
    } catch (e) {
        console.error("Cache Error at tg_ guard:", e);
    }

    if (data.callback_query) {
        const cq = data.callback_query;
        const msgId = cq.message ? cq.message.message_id : "";
        const chatId = cq.message ? cq.message.chat.id : "";
        const lockKey = `lock_cb_${chatId}_${msgId}`;

        let isLocked = false;
        try {
            isLocked = cache.get(lockKey);
        } catch (e) {
            console.error("Cache Error at debounce get:", e);
        }

        if (isLocked) {
            console.warn(`[Debounce Guard] Menolak klik beruntun pada pesan ${msgId}.`);
            try {
                UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEBOT_TOKEN}/answerCallbackQuery`, {
                    method: 'post',
                    contentType: 'application/json',
                    muteHttpExceptions: true,
                    payload: JSON.stringify({
                        callback_query_id: cq.id,
                        text: "⏳ Sedang diproses, mohon tunggu sebentar...",
                        show_alert: false
                    })
                });
            } catch (e) { /* abaikan */ }
            return HtmlService.createHtmlOutput("OK");
        }

        // Pasang lock selama 6 detik untuk menahan request ganda
        try {
            cache.put(lockKey, "1", 6);
        } catch (e) {
            console.error("Cache Error at debounce put:", e);
        }

        try {
            UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEBOT_TOKEN}/answerCallbackQuery`, {
                method: 'post',
                contentType: 'application/json',
                muteHttpExceptions: true,
                payload: JSON.stringify({ callback_query_id: cq.id, text: "Memproses..." })
            });
        } catch (e) { /* abaikan */ }
    }

    prosesPendingUpdate(updateId, data);

    return HtmlService.createHtmlOutput("OK");
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
        return HtmlService.createHtmlOutput("OK");
    }

    prosesPendingWhatsAppUpdate(waUniqueId, data);

    return HtmlService.createHtmlOutput("OK");
}

function prosesPencarianTiket(userPhone, ticketId) {
    const cleanTicketId = (ticketId || "").trim();

    if (!isValidTicketFormat(cleanTicketId)) {
        kirimWaTeksBiasa(userPhone, `⚠️ *Format Tiket Tidak Valid*\n\nFormat nomor tiket yang Anda masukkan tidak sesuai.\nContoh format tiket yang benar: *PMD-20260913-2-ABCD*\n\nSilakan periksa kembali karakter dan ejaan tiket Anda.`);
        return;
    }

    const cache = CacheService.getScriptCache();
    const rateKey = `rate_wa_search_${userPhone}`;
    if (cache.get(rateKey)) {
        kirimWaTeksBiasa(userPhone, `⏳ Mohon tunggu beberapa detik sebelum melakukan pencarian tiket kembali.`);
        return;
    }
    try {
        cache.put(rateKey, "1", 3);
    } catch (e) { /* abaikan error cache */ }

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const data = sheet.getDataRange().getValues();

    const colIndexTiket = 9;
    const colIndexStatus = 10;
    const colIndexNama = 1;
    const colIndexUlasan = 12;
    const colIndexBidangPJ = 11;
    const colIndexKeperluan = 7;

    let rowFound = -1;

    for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][colIndexTiket].toString().toUpperCase() === cleanTicketId.toUpperCase()) {
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
            `🎫 *ID Tiket:* ${cleanTicketId}\n` +
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
        kirimWaTeksBiasa(userPhone, `❌ *Tiket Tidak Ditemukan*\n\nMaaf, nomor tiket *${cleanTicketId}* tidak terdaftar di sistem kami. Mohon periksa kembali karakter dan ejaan tiket Anda.`);
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
        const changeValue = data.entry?.[0]?.changes?.[0]?.value;

        if (Array.isArray(changeValue?.messages) && changeValue.messages.length > 0) {
            const messageObject = changeValue.messages[0];
            const userPhone = messageObject.from;

            if (messageObject.type === "text" && messageObject.text?.body) {
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

function detectIsDev(data) {
    const isCallback = !!data.callback_query;
    const isReply = !isCallback && data.message && data.message.reply_to_message;
    if (isCallback && data.callback_query.data && data.callback_query.data.startsWith("DEV_")) return true;
    if (isReply) {
        const repliedText = data.message.reply_to_message.text || "";
        if (repliedText.includes("_DEV")) return true;
    }
    return false;
}

function prosesPendingUpdate(updateId, data) {
    const startTime = new Date();
    let statusLog = "START";
    let detailLog = "";

    try {
        if (!data) return;

        const isCallback = !!data.callback_query;
        const isReply = !isCallback && data.message && data.message.reply_to_message;

        let isDev = detectIsDev(data);

        if (!isDev) {
            let msgText = "";
            if (isCallback && data.callback_query.message) {
                msgText = data.callback_query.message.text || "";
            } else if (isReply && data.message.reply_to_message) {
                msgText = data.message.reply_to_message.text || "";
            }
            if (msgText.includes("[TEST]") || msgText.includes("DEV-")) {
                isDev = true;
                console.log("[Lightweight DEV Detection] Teks pesan mengandung marker DEV.");
            }
        }

        let _sheet = null;
        function getSheet() {
            if (!_sheet) {
                const targetSpreadsheetId = isDev ? DEV_SPREADSHEET_ID : SPREADSHEET_ID;
                const ss = SpreadsheetApp.openById(targetSpreadsheetId);
                _sheet = ss.getSheetByName(NAMA_TAB_UTAMA);
                if (!_sheet) throw new Error("Sheet utama tidak ditemukan");
            }
            return _sheet;
        }

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
            let callbackData = cq.data;
            if (callbackData.startsWith("DEV_")) {
                callbackData = callbackData.substring(4);
            }
            const parts = callbackData.split('_');
            const aksiLog = parts[0];
            const rowLog = parseInt(parts[parts.length - 1]);

            console.log(`[CALLBACK] Processing callbackData: ${cq.data} (isDev: ${isDev})`);

            if (isNaN(rowLog) || !validateRowBounds(getSheet(), rowLog)) throw new Error("Row tidak valid atau di luar jangkauan spreadsheet");

            // ==================== BATAL ====================
            if (aksiLog === "BATAL") {
                // Hapus lock agar tombol bidang bisa diklik kembali jika dibatalkan
                const cache = CacheService.getScriptCache();
                const msgId = cq.message ? cq.message.message_id : "";
                const chatId = cq.message ? cq.message.chat.id : "";
                cache.remove(`lock_cb_${chatId}_${msgId}`);

                const rawText = cq.message ? cq.message.text || "" : "";
                const cleanText = rawText.split(/\n\n(?:[✅⏳]\s*)?\*?STATUS:/i)[0].trim();

                getSheet().getRange(rowLog, 11, 1, 2).setValues([["Menunggu Tindakan", ""]]);
                UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEBOT_TOKEN}/editMessageText`, {
                    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
                    payload: JSON.stringify({
                        chat_id: cq.message.chat.id,
                        message_id: cq.message.message_id,
                        text: cleanText,
                        parse_mode: 'Markdown',
                        reply_markup: default_keyboard(rowLog, isDev)
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
                const rawText = cq.message ? cq.message.text || "" : "";

                // GUARD: Jika pesan sudah berstatus 'Selesai' atau 'Ditolak', skip!
                if (rawText.includes("STATUS: Selesai") || rawText.includes("STATUS: Ditolak")) {
                    console.warn(`[PROSES Guard] Pesan ${cq.message.message_id} sudah final. Skip.`);
                    statusLog = "SKIPPED_ALREADY_PROCESSED";
                    detailLog = "Laporan sudah berstatus Selesai atau Ditolak";
                    return;
                }

                const baseText = rawText.split(/\n\n(?:[✅⏳]\s*)?\*?STATUS:/i)[0].trim();
                const pesanUpdate = `${baseText}\n\n✅ *STATUS: Diproses*\nOleh: ${namaBidang}\nPetugas: ${cq.from.first_name}`;
                const devPrefix = isDev ? "DEV_" : "";

                const keyboard = {
                    inline_keyboard: [
                        [{ text: "✅ Tandai Selesai", callback_data: `${devPrefix}TANYA_SELESAI_${rowLog}` }],
                        [{ text: "❌ Batal", callback_data: `${devPrefix}BATAL_${rowLog}` }]
                    ]
                };

                // Update tampilan Telegram TERLEBIH DAHULU agar respon terasa instan di UI admin (<300ms)
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

                // Tulis ke Spreadsheet setelah tampilan Telegram ter-update
                getSheet().getRange(rowLog, 11, 1, 2).setValues([["Diproses", namaBidang]]);

                statusLog = "SUCCESS";
                detailLog = `Diproses oleh ${namaBidang}`;
            }

            // ==================== TANYA (Tolak / Selesai) ====================
            else if (aksiLog === "TANYA") {
                const jenis = parts[1]; // "SELESAI" atau "TOLAK"
                const rowStr = parts[2];

                let idTicket = "Unknown";
                const rawText = cq.message ? cq.message.text || "" : "";
                const baseText = rawText.split(/\n\n(?:[✅⏳]\s*)?\*?STATUS:/i)[0].trim();

                const matchTicket = baseText.match(/ID Tiket:\s*([^\s\n]+)/i);
                if (matchTicket) {
                    idTicket = matchTicket[1].trim();
                }

                const kataTugas = jenis === "TOLAK" ? "alasan penolakan" : "catatan penyelesaian";
                const username = cq.from.username ? "@" + cq.from.username : cq.from.first_name;
                const originalMessageId = cq.message.message_id;

                const labelStatus = jenis === "TOLAK" ? "Menunggu Alasan Penolakan" : "Menunggu Catatan Penyelesaian";
                const pesanUpdate = `${baseText}\n\n⏳ *STATUS: ${labelStatus}*\nOleh: ${username}`;
                const devPrefix = isDev ? "DEV_" : "";

                const keyboard = {
                    inline_keyboard: [
                        [{ text: "❌ Batal", callback_data: `${devPrefix}BATAL_${rowLog}` }]
                    ]
                };

                const devTag = isDev ? "_DEV" : "";
                const replyMsg = `[#ROW_${rowStr}_${jenis}${devTag}] [#MAIN_${originalMessageId}] ${username}, silakan balas (reply) pesan ini dengan *${kataTugas}* untuk laporan ini.\nTiket: ${idTicket}`;

                const telegramBaseUrl = `https://api.telegram.org/bot${TELEBOT_TOKEN}`;

                // Batching paralel request ke Telegram (edit pesan + sendMessage instruksi reply)
                UrlFetchApp.fetchAll([
                    {
                        url: `${telegramBaseUrl}/editMessageText`,
                        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
                        payload: JSON.stringify({
                            chat_id: cq.message.chat.id,
                            message_id: cq.message.message_id,
                            text: pesanUpdate,
                            parse_mode: 'Markdown',
                            reply_markup: keyboard
                        })
                    },
                    {
                        url: `${telegramBaseUrl}/sendMessage`,
                        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
                        payload: JSON.stringify({
                            chat_id: TELEGROUP_ID,
                            text: replyMsg,
                            parse_mode: 'Markdown',
                            reply_to_message_id: originalMessageId,
                            reply_markup: { force_reply: true, selective: true }
                        })
                    }
                ]);

                statusLog = "SUCCESS";
                detailLog = `Instruksi ${jenis} dikirim | MAIN_ID: ${originalMessageId}`;
            }
        }

        // ==================== REPLY MESSAGE ====================
        else if (isReply) {
            const repliedText = data.message.reply_to_message ? data.message.reply_to_message.text || "" : "";
            const matchRow = repliedText.match(/#ROW_(\d+)_([A-Z]+)(_DEV)?/);
            const matchMain = repliedText.match(/#MAIN_(\d+)/);

            if (matchRow && matchMain) {
                rowLog = parseInt(matchRow[1]);
                const jenis = matchRow[2];
                if (matchRow[3] === "_DEV") {
                    isDev = true;
                }
                aksiLog = jenis;
                const originalMessageId = parseInt(matchMain[1]);

                const sheet = getSheet();
                if (isNaN(rowLog) || !validateRowBounds(sheet, rowLog)) {
                    throw new Error("Row reply tidak valid atau di luar jangkauan spreadsheet");
                }

                const teksBalasan = data.message.text || "";
                const safeTeksBalasan = sanitizeInputText(teksBalasan, 1000);
                const pengisi = data.message.from.first_name || "Petugas";
                const chatId = data.message.chat.id;

                // Batch read: ambil kolom J(10), K(11), L(12), O(15) sekaligus dalam 1 panggilan
                const rowData = sheet.getRange(rowLog, 1, 1, 15).getValues()[0];
                const idTicket = rowData[9] || "";   // Kolom J (10)
                const currentStatus = (rowData[10] || "").toString().trim().toUpperCase(); // Kolom K (11)
                const bidangPJ = rowData[11] || "";  // Kolom L (12)
                const noWaRaw = rowData[3] || "";    // Kolom D (4)
                const telegramMsgIdInSheet = rowData[14] || ""; // Kolom O (15)

                // IDEMPOTENCY GUARD
                if (currentStatus === "SELESAI" || currentStatus === "DITOLAK") {
                    console.warn(`[REPLY Guard] Baris ${rowLog} sudah '${currentStatus}'. Skip.`);
                    statusLog = "SKIPPED_ALREADY_PROCESSED";
                    detailLog = `Laporan sudah berstatus ${currentStatus}`;
                    return;
                }

                console.log(`[REPLY] Processing ${jenis} for row ${rowLog}, ticket: ${idTicket} (isDev: ${isDev})`);

                const statusBaru = jenis === "TOLAK" ? "Ditolak" : "Selesai";

                // Batch write status (11), bidang (12), dan catatan (13) sekaligus dalam 1 panggilan API
                sheet.getRange(rowLog, 11, 1, 3).setValues([[statusBaru, bidangPJ, safeTeksBalasan]]);

                const requests = [];

                // Push WhatsApp API call ke array request paralel jika SELESAI
                if (jenis === "SELESAI" && noWaRaw && noWaRaw !== "KOSONG" && noWaRaw !== "") {
                    const noWaBersih = formatNomorWA(noWaRaw);
                    const pesanSelesai =
                        `Halo Bapak/Ibu,\n\nKami informasikan bahwa laporan/pengaduan Anda dengan nomor tiket *${idTicket}* telah *SELESAI* ditindaklanjuti.\n\nSebagai upaya peningkatan kualitas pelayanan publik, silakan Bapak/Ibu mengisi Survei Kepuasan Masyarakat (SKM) melalui tautan berikut:\n🔗 ${SKM_URL}\n\nSetelah mengisi survei tersebut, Anda dapat melihat detail penyelesaian dengan mengetik:\n*CEK ${idTicket}*\n\nTerima kasih atas partisipasi Anda telah menggunakan SI Buku Tamu Online DPMD Kabupaten Muara Enim. 🙏`;

                    requests.push({
                        url: `https://graph.facebook.com/v25.0/${WA_PHONE_NUMBER_ID}/messages`,
                        method: 'post',
                        contentType: 'application/json',
                        headers: { "Authorization": "Bearer " + WA_ACCESS_TOKEN },
                        payload: JSON.stringify({
                            "messaging_product": "whatsapp",
                            "recipient_type": "individual",
                            "to": noWaBersih,
                            "type": "text",
                            "text": { "preview_url": true, "body": pesanSelesai }
                        }),
                        muteHttpExceptions: true
                    });
                }

                // === Batch Telegram API calls ===
                const targetMsgId = telegramMsgIdInSheet || originalMessageId;
                const statusUpper = statusBaru.toUpperCase();
                const safeIdTicketEsc = escapeMarkdown(idTicket);
                const safeTeksBalasanEsc = escapeMarkdown(safeTeksBalasan);
                const safeBidangEsc = escapeMarkdown(bidangPJ);
                const safePengisiEsc = escapeMarkdown(pengisi);

                let syncText = "";
                if (statusUpper === "DITOLAK") {
                    syncText = `❌ *LAPORAN DITOLAK*\n🎫 Tiket: ${safeIdTicketEsc}\n📝 Alasan: ${safeTeksBalasanEsc}\n\n_Pesan ini telah difinalisasi._`;
                } else if (statusUpper === "SELESAI") {
                    syncText = `✅ *LAPORAN SELESAI*\n🎫 Tiket: ${safeIdTicketEsc}\n🏢 Bidang: ${safeBidangEsc}\n📝 Catatan: ${safeTeksBalasanEsc}\n\n_Pesan ini telah difinalisasi._`;
                }

                const telegramBaseUrl = `https://api.telegram.org/bot${TELEBOT_TOKEN}`;

                // 1. Update pesan utama
                if (syncText) {
                    requests.push({
                        url: `${telegramBaseUrl}/editMessageText`,
                        method: 'post',
                        contentType: 'application/json',
                        muteHttpExceptions: true,
                        payload: JSON.stringify({
                            chat_id: TELEGROUP_ID,
                            message_id: targetMsgId,
                            text: syncText,
                            parse_mode: 'Markdown',
                            reply_markup: { inline_keyboard: [] }
                        })
                    });
                }

                // 2. Sembunyikan markup pesan reply instruksi
                requests.push({
                    url: `${telegramBaseUrl}/editMessageReplyMarkup`,
                    method: 'post',
                    contentType: 'application/json',
                    muteHttpExceptions: true,
                    payload: JSON.stringify({
                        chat_id: chatId,
                        message_id: originalMessageId,
                        reply_markup: { inline_keyboard: [] }
                    })
                });

                // 3. Kirim konfirmasi ke grup
                requests.push({
                    url: `${telegramBaseUrl}/sendMessage`,
                    method: 'post',
                    contentType: 'application/json',
                    muteHttpExceptions: true,
                    payload: JSON.stringify({
                        chat_id: chatId,
                        text: `✅ Laporan ${safeIdTicketEsc} (${isDev ? 'DEV' : 'PROD'}) berhasil ${jenis === "TOLAK" ? "DITOLAK" : "DISELESAIKAN"} oleh ${safePengisiEsc}.`,
                        parse_mode: "Markdown"
                    })
                });

                // Kirim SEMUA request network (WhatsApp + Telegram) secara PARALEL dalam 1 batch
                UrlFetchApp.fetchAll(requests);

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
        console.log(`[prosesPendingUpdate] updateId=${updateId} status=${statusLog} duration=${duration}ms detail=${detailLog}`);
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
    const url = `https://api.telegram.org/bot${TELEBOT_TOKEN}/getWebhookInfo`;
    const response = UrlFetchApp.fetch(url);
    console.log(response.getContentText());
}

function _getExecUrl() {
    let url = WEBAPP_EXEC_URL;
    if (!url || url === "") {
        url = ScriptApp.getService().getUrl();
    }
    if (url.endsWith("/dev")) {
        throw new Error(
            "❌ GAGAL: URL webhook mengarah ke /dev (development)!\n" +
            "Telegram TIDAK BISA mengakses URL /dev.\n\n" +
            "SOLUSI:\n" +
            "1. Buka Deploy > Manage deployments\n" +
            "2. Salin URL deployment yang berakhiran /exec\n" +
            "3. Tempel ke konstanta WEBAPP_EXEC_URL di baris atas kode\n" +
            "4. Simpan & jalankan ulang fungsi ini"
        );
    }
    return url;
}

function activateTelegramWebhook() {
    const webhookUrl = _getExecUrl();
    const url = `https://api.telegram.org/bot${TELEBOT_TOKEN}/setWebhook`;
    const response = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
            url: webhookUrl,
            drop_pending_updates: true
        })
    });
    console.log("Activate webhook result: " + response.getContentText());
    console.log("Webhook URL: " + webhookUrl);
    webhookTelegramCheck();
}

function resetPendingUpdates() {
    const webhookUrl = _getExecUrl();
    const url = `https://api.telegram.org/bot${TELEBOT_TOKEN}/setWebhook`;
    const response = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
            url: webhookUrl,
            drop_pending_updates: true
        })
    });
    console.log("Reset pending updates: " + response.getContentText());
    webhookTelegramCheck();
}

function cleanupOldProperties() {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    let deleted = 0;
    for (const key in allProps) {
        if (key.startsWith("processed_") || key.startsWith("pending_") || key.startsWith("processed_tg_")) {
            props.deleteProperty(key);
            deleted++;
        }
    }
    console.log(`Berhasil membersihkan ${deleted} properties lama.`);
}

function deactivateTelegramWebhook() {
    const url = `https://api.telegram.org/bot${TELEBOT_TOKEN}/setWebhook?url=`;
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

function onSpreadsheetEdit(e) {
    try {
        if (!e || !e.range) return;

        const sheet = e.range.getSheet();
        if (sheet.getName() !== NAMA_TAB_UTAMA) return;

        const row = e.range.getRow();
        if (row < 2) return; // Abaikan baris header
        if (!validateRowBounds(sheet, row)) return;

        const col = e.range.getColumn();
        // Hanya proses jika yang diedit adalah Kolom K, L, atau M
        if (col !== 11 && col !== 12 && col !== 13) return;

        const spreadsheetId = sheet.getParent().getId();
        const isDev = (spreadsheetId === DEV_SPREADSHEET_ID);

        let status = (sheet.getRange(row, 11).getValue() || "").toString().trim();
        let bidang = (sheet.getRange(row, 12).getValue() || "").toString().trim();
        let catatan = (sheet.getRange(row, 13).getValue() || "").toString().trim();
        const idTicket = (sheet.getRange(row, 10).getValue() || "").toString().trim();
        const telegramMsgId = sheet.getRange(row, 15).getValue(); // Kolom O (15)

        console.log(`[onSpreadsheetEdit] Edit di ${isDev ? 'DEV' : 'PROD'} Baris ${row} Kolom ${col}. Tiket: ${idTicket}, Status: ${status}`);

        // ========== GUARD: Abaikan jika ID Tiket tidak ada (bukan baris data valid) ==========
        if (!idTicket) return;

        // ========== AUTO-CORRECTION: Menunggu + Bidang Diisi → Advance ke Diproses ==========
        if ((status.toUpperCase() === "MENUNGGU TINDAKAN" || status === "") && bidang !== "") {
            status = "Diproses";
            sheet.getRange(row, 11).setValue("Diproses");
            console.log(`[Auto-Correction Row ${row}] Status di-advance ke 'Diproses' karena Bidang PJ diisi.`);
        }

        // ========== SINKRONISASI TELEGRAM ==========
        if (!telegramMsgId) {
            console.warn(`[onSpreadsheetEdit Row ${row}] Kolom O (Telegram_Message_ID) kosong. Sinkronisasi Telegram dilewati.`);
        } else {
            syncEditToTelegram(telegramMsgId, status, bidang, catatan, idTicket, row, isDev);
        }

        // ========== SAFETY GUARD: Kirim WA Selesai hanya jika Catatan sudah diisi ==========
        if (status.toUpperCase() === "SELESAI") {
            if (catatan !== "") {
                const noWaRaw = sheet.getRange(row, 4).getValue(); // Kolom D (4) = Nomor WA
                if (noWaRaw && noWaRaw !== "KOSONG" && noWaRaw !== "") {
                    const noWaBersih = formatNomorWA(noWaRaw.toString());

                    // Cache agar WA Selesai tidak dikirim berulang untuk tiket yang sama
                    const cache = CacheService.getScriptCache();
                    const cacheKey = `wa_selesai_${isDev ? 'dev_' : ''}${row}_${idTicket}`;
                    if (!cache.get(cacheKey)) {
                        const pesanSelesai =
                            `Halo Bapak/Ibu,\n\nKami informasikan bahwa laporan/pengaduan Anda dengan nomor tiket *${idTicket}* telah *SELESAI* ditindaklanjuti.\n\nCatatan Penyelesaian:\n${catatan}\n\nSebagai upaya peningkatan kualitas pelayanan publik, silakan Bapak/Ibu mengisi Survei Kepuasan Masyarakat (SKM) melalui tautan berikut:\n🔗 ${SKM_URL}\n\nSetelah mengisi survei tersebut, Anda dapat melihat detail penyelesaian dengan mengetik:\n*CEK ${idTicket}*\n\nTerima kasih atas partisipasi Anda telah menggunakan SI Buku Tamu Online DPMD Kabupaten Muara Enim. 🙏`;

                        kirimWaTeksBiasa(noWaBersih, pesanSelesai);
                        cache.put(cacheKey, "true", 86400); // 24 jam agar tidak dikirim ulang
                        console.log(`[onSpreadsheetEdit Row ${row}] WA Selesai berhasil dikirim ke ${noWaBersih}`);
                    }
                }
            } else {
                console.log(`[Safety Guard Row ${row}] Status = Selesai tetapi Catatan (Kolom M) masih kosong. WA DITAHAN.`);
            }
        }

    } catch (err) {
        console.error("Error pada onSpreadsheetEdit: " + err.toString());
    }
}

function syncEditToTelegram(messageId, status, bidang, catatan, idTicket, row, isDev = false) {
    const statusUpper = status.toUpperCase();
    let text = "";
    let keyboard = { inline_keyboard: [] };
    const devPrefix = isDev ? "DEV_" : "";

    const safeTicket = escapeMarkdown(idTicket);
    const safeBidang = escapeMarkdown(sanitizeInputText(bidang, 100));
    const safeCatatan = escapeMarkdown(sanitizeInputText(catatan, 1000));
    const safeStatus = escapeMarkdown(status);

    if (statusUpper === "MENUNGGU TINDAKAN" || statusUpper === "") {
        text = `🚨 *Laporan/Tamu Baru Masuk!* 🚨\n\n📌 *ID Tiket:* ${safeTicket}\n\nSilakan tentukan bidang penanggung jawab:`;
        keyboard = default_keyboard(row, isDev);

    } else if (statusUpper === "DIPROSES") {
        const bidangStr = safeBidang || "(Bidang belum ditentukan)";
        text = `🚨 *Laporan/Tamu Baru Masuk!* 🚨\n\n📌 *ID Tiket:* ${safeTicket}\n\n✅ *STATUS: Diproses*\nOleh: ${bidangStr}\n_(Diperbarui dari Spreadsheet${isDev ? ' DEV' : ''})_`;
        keyboard = {
            inline_keyboard: [
                [{ text: "✅ Tandai Selesai", callback_data: `${devPrefix}TANYA_SELESAI_${row}` }],
                [{ text: "❌ Batal", callback_data: `${devPrefix}BATAL_${row}` }]
            ]
        };

    } else if (statusUpper === "SELESAI") {
        const bidangStr = safeBidang || "(Bidang)";
        const catatanStr = safeCatatan || "(Catatan belum diisi)";
        text = `🚨 *Laporan/Tamu Baru Masuk!* 🚨\n\n📌 *ID Tiket:* ${safeTicket}\n\n🎉 *STATUS: Selesai*\nOleh: ${bidangStr}\nCatatan: ${catatanStr}\n_(Diperbarui dari Spreadsheet${isDev ? ' DEV' : ''})_`;
        keyboard = { inline_keyboard: [] }; // Hapus semua tombol

    } else if (statusUpper === "DITOLAK") {
        const alasanStr = safeCatatan || "(Alasan belum diisi)";
        text = `🚨 *Laporan/Tamu Baru Masuk!* 🚨\n\n📌 *ID Tiket:* ${safeTicket}\n\n❌ *STATUS: Ditolak*\nAlasan: ${alasanStr}\n_(Diperbarui dari Spreadsheet${isDev ? ' DEV' : ''})_`;
        keyboard = { inline_keyboard: [] }; // Hapus semua tombol

    } else {
        // Status tidak dikenal — tetap update teks apa adanya
        text = `🚨 *Laporan/Tamu Baru Masuk!* 🚨\n\n📌 *ID Tiket:* ${safeTicket}\n\n📊 *Status:* ${safeStatus}\n_(Diperbarui dari Spreadsheet${isDev ? ' DEV' : ''})_`;
        keyboard = { inline_keyboard: [] };
    }

    const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEBOT_TOKEN}/editMessageText`, {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: true,
        payload: JSON.stringify({
            chat_id: TELEGROUP_ID,
            message_id: messageId,
            text: text,
            parse_mode: 'Markdown',
            reply_markup: keyboard
        })
    });
    console.log(`[syncEditToTelegram] Response Telegram untuk MsgID ${messageId}: ${res.getResponseCode()}`);
}

function setupOnEditTrigger() {
    // 1. Pasang Trigger untuk Spreadsheet Produksi
    try {
        const ssProd = SpreadsheetApp.openById(SPREADSHEET_ID);
        ScriptApp.newTrigger('onSpreadsheetEdit')
            .forSpreadsheet(ssProd)
            .onEdit()
            .create();
        Logger.log("Trigger onSpreadsheetEdit berhasil dipasang untuk PROD!");
    } catch (e) {
        Logger.log("Gagal memasang trigger PROD: " + e.toString());
    }

    // 2. Pasang Trigger untuk Spreadsheet DEV
    try {
        const ssDev = SpreadsheetApp.openById(DEV_SPREADSHEET_ID);
        ScriptApp.newTrigger('onSpreadsheetEdit')
            .forSpreadsheet(ssDev)
            .onEdit()
            .create();
        Logger.log("Trigger onSpreadsheetEdit berhasil dipasang untuk DEV!");
    } catch (e) {
        Logger.log("Gagal memasang trigger DEV: " + e.toString());
    }
}