const CONFIG = {
    DRIVE_FOLDER_ID: "1QCFpQb9XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", // 雲端硬碟資料夾 FOLDER
    SPREADSHEET_ID: "1CG9xXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", // 試算表 ID
    ADMIN_EMAIL: "notify.admin@gmail.com", // 管理員郵件
    VENDOR_EMAIL: "notify.vendor@gmail.com.tw", // 廠商郵件（可替換）
    NOTIFICATION_EMAIL: "notify.bpm@gmail.com", // 通知郵件 - 啟用問卷責任單位
    NOTIFY_QA_EMAIL:"notify.QA@gmail.com" // 品保通知郵件（可替換）
};

/**
 * 記錄錯誤並通知管理員
 */
function handleError(functionName, error, context = "") {
    const message = `${functionName} 錯誤: ${error.message}\n上下文: ${context}`;
    Logger.log(message);
    try {
        GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, `${functionName} 錯誤`, message);
        Logger.log(`錯誤通知已寄送至: ${CONFIG.ADMIN_EMAIL}`);
    } catch (emailError) {
        Logger.log(`寄送錯誤通知失敗: ${emailError.message}`);
    }
}

/**
 * 處理 POST 請求
 */
function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);

        // ← 從 POST 傳入
        const { ticketNo, materialName, descript, imageUrl, aplyDate, abnItem, vendorEmail, notificationEmail, notifyQaEmail, vendorName } = data;

        // 使用傳入的 email，如果未提供則預設為空字串或用 ADMIN_EMAIL
        const VENDOR_EMAIL = vendorEmail || CONFIG.ADMIN_EMAIL;
        const NOTIFICATION_EMAIL = notificationEmail || CONFIG.ADMIN_EMAIL;
        const NOTIFY_QA_EMAIL = notifyQaEmail || CONFIG.NOTIFY_QA_EMAIL;

        // 驗證必要欄位!vendorName
        if (!ticketNo || !materialName || !descript ||   !vendorEmail) {
            throw new Error("缺少必要欄位:單號、料名 或 異常描述、廠商email");
        }

        // 創建 Google 表單
        const form = createForm(ticketNo, materialName, descript, imageUrl, VENDOR_EMAIL, NOTIFICATION_EMAIL,NOTIFY_QA_EMAIL, vendorName);

        // 寫入試算表 - 採購紀記錄
        const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName("表單紀錄") || spreadsheet.insertSheet("表單紀錄");

        //插入試算表的順序											
        sheet.appendRow([
            ticketNo,
            aplyDate || "",
            abnItem || "",
            materialName,
            descript,  
            vendorName  || "",
            imageUrl || "",
            form.getPublishedUrl(),
            new Date(),
            VENDOR_EMAIL,
            NOTIFICATION_EMAIL,
            NOTIFY_QA_EMAIL
        ]);
         

        Logger.log(`表單紀錄寫入成功: ticketNo=${ticketNo}`);

        return ContentService.createTextOutput(
            JSON.stringify({
                status: "success",
                formUrl: form.getPublishedUrl(),
            })
        ).setMimeType(ContentService.MimeType.JSON);

    } catch (error) {
        handleError("doPost", error, `POST 資料: ${e.postData?.contents}`);
        return ContentService.createTextOutput(
            JSON.stringify({
                status: "error",
                message: error.message,
            })
        ).setMimeType(ContentService.MimeType.JSON);
    }
}

/**
 * 創建 Google 表單
 */
function createForm(ticketNo, materialName, descript, imageUrl, VENDOR_EMAIL, NOTIFICATION_EMAIL,NOTIFY_QA_EMAIL, vendorName) {
    // 單號：ticketNo , 
    // 料名：materialName, 
    // 異常描述：descript, 
    // 照片路徑：imageUrl, 
    // 廠商MAIL：VENDOR_EMAIL, 
    // QA品保MAIL：NOTIFICATION_EMAIL,
    // 採購MAIL：NOTIFY_QA_EMAIL
    // 廠商名稱 VENDOR_NAME
    const vendorNameSafe = vendorName || "";
    try {
        // 強制創建新表單
        const formTitle = `產品異常回覆單 - ${ticketNo}`;
        let form = FormApp.create(formTitle);
        Logger.log(`創建新表單: ${formTitle}`);
 
        form.setTitle(`產品異常回覆單 ${ticketNo}`);
        form.setDescription(`
            廠商名稱: ${vendorNameSafe}
            品名: ${materialName}
            異常描述: ${descript}
            廠商EMAIL: ${VENDOR_EMAIL}
            表單聯絡人EMAIL: ${NOTIFICATION_EMAIL}
            📢 注意事項 / Notice:
            - 將附件寄至 ${NOTIFY_QA_EMAIL}（副本  ${NOTIFICATION_EMAIL}，主旨包含單號 ${ticketNo}），或上傳至雲端並提供連結。           

            如有問題，請聯繫管理員：${CONFIG.ADMIN_EMAIL}
            感謝您的配合！
        `);

         // 限制單次提交
        form.setLimitOneResponsePerUser(true);

        // 添加異常圖片
        if (imageUrl) {
            var baseUrl = "https://bpm.company.com.tw/";
            var imageUrlsArray = imageUrl
                .split(";")
                .filter(Boolean)
                .map((path) => path.replace("../", baseUrl));
            Logger.log(imageUrlsArray);

            imageUrlsArray.forEach((url, index) => {
                try {
                    const blob = UrlFetchApp.fetch(url).getBlob();
                    form.addImageItem()
                        .setImage(blob)
                        .setTitle(`異常圖片 ${index + 1}`);
                } catch (e) {
                    form.addTextItem().setTitle("異常圖片").setHelpText(`圖片載入失敗：${url}`);
                    Logger.log(`圖片載入錯誤: ${e.message}`);
                }
            });
        }

        // 添加表單欄位
        // form.addTextItem().setTitle("廠商名稱/ Vendor Name").setRequired(true);
        form.addParagraphTextItem().setTitle("異常處理說明/ Abnormality Handling Description");
        form.addTextItem().setTitle("聯絡人姓名/ Contact Person Name").setRequired(true);
        form.addTextItem().setTitle("聯絡電話/ Contact Phone");
        form.addParagraphTextItem().setTitle("應急措施/ Emergency Measures");
        form.addParagraphTextItem().setTitle("原因分析/ Root Cause Analysis").setRequired(true);
        form.addParagraphTextItem().setTitle("矯正措施/ Corrective Actions").setRequired(true);
        form.addParagraphTextItem().setTitle("預防措施/ Preventive Actions").setRequired(true);
        form.addParagraphTextItem().setTitle("附件連結");


        // 清理所有 onFormSubmit 觸發器
        ScriptApp.getProjectTriggers().forEach((trigger) => {
            if (trigger.getHandlerFunction() === "onFormSubmit") {
                ScriptApp.deleteTrigger(trigger);
                Logger.log(`已刪除舊觸發器: ${trigger.getUniqueId()}`);
            }
        });

        // 創建新觸發器
        ScriptApp.newTrigger("onFormSubmit").forForm(form).onFormSubmit().create();
        Logger.log(`新觸發器創建 for 表單: ${form.getId()}`);

        // 寄送表單連結
        sendFormEmail(VENDOR_EMAIL, ticketNo, materialName, descript, form.getPublishedUrl(), vendorName, NOTIFICATION_EMAIL, NOTIFY_QA_EMAIL);

        return form;
    } catch (error) {
        handleError("createForm", error, `ticketNo: ${ticketNo}`);
        throw error;
    }
}

/**
 * 處理所有表單提交（防止重複）
 */
function onFormSubmit(e) {
    try {
        const formResponse = e.response;
        const responseId = formResponse.getId(); // 回應唯一 ID
        const form = e.source;
        const formTitle = form.getTitle();
        const ticketNo = formTitle.replace("產品異常回覆單 ", "").trim();

        // 檢查是否已處理該回應
        const properties = PropertiesService.getScriptProperties();
        if (properties.getProperty(responseId)) {
            Logger.log(`回應已處理，跳過: responseId=${responseId}, ticketNo=${ticketNo}`);
            return;
        }

        // 解析描述中的品名、廠商名稱、mail
        const description = form.getDescription().split("\n");
        const materialLine = description.find((line) => line.trim().startsWith("品名:"));
        const materialName = materialLine ? materialLine.replace("品名:", "").trim() : "未知";
        const vendorLine = description.find((line) => line.trim().startsWith("廠商名稱:"));
        const vendorName = vendorLine ? vendorLine.replace("廠商名稱:", "").trim() : "未知";
        const vendorEmailstr = description.find((line) => line.trim().startsWith("廠商EMAIL:"));
        const vendorEmail = vendorEmailstr ? vendorEmailstr.replace("廠商EMAIL:", "").trim() : "未知";
        const notificationstr = description.find((line) => line.trim().startsWith("表單聯絡人EMAIL:"));
        const notificationEmail = notificationstr ? notificationstr.replace("表單聯絡人EMAIL:", "").trim() : "未知";

        Logger.log(`解析 email 完成: vendorEmail=${vendorEmail}, notificationEmail=${notificationEmail}`);

        const itemResponses = formResponse.getItemResponses();
        Logger.log(`表單提交: ticketNo=${ticketNo}, responseId=${responseId}, 回應數=${itemResponses.length}`);

        // 動態映射回應
        const vendorData = { ticketNo, vendorName, submitTime: formResponse.getTimestamp(), materialName };
        let attachmentLink = "";

        //處理問卷回應資料廠商名稱  異常處理說明 聯絡電話 應急措施" 改為非必填
        itemResponses.forEach((response) => {
            const title = response.getItem().getTitle();
            const titleName = title.split("/")[0].trim();
            // if (titleName === "廠商名稱") vendorData.vendorName = response.getResponse();
            if (titleName === "異常處理說明") vendorData.handlingDesc = response.getResponse();
            else if (titleName === "聯絡人姓名") vendorData.contactName = response.getResponse();
            else if (titleName === "聯絡電話") vendorData.contactPhone = response.getResponse();
            else if (titleName === "應急措施") vendorData.reply1 = response.getResponse();
            else if (titleName === "原因分析") vendorData.reply2 = response.getResponse();
            else if (titleName === "矯正措施") vendorData.reply3 = response.getResponse();
            else if (titleName === "預防措施") vendorData.reply4 = response.getResponse();
            else if (titleName === "附件連結") {
                if (!attachmentLink) attachmentLink = response.getResponse();
            }
        });

        // 驗證必要欄位
        if (
            !vendorData.contactName ||         
            !vendorData.reply2 ||
            !vendorData.reply3 ||
            !vendorData.reply4
        ) {
            throw new Error("請填寫所有必填欄位資料");
        }

        // 檢查試算表是否已有相同回應
        const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName("廠商回應") || spreadsheet.insertSheet("廠商回應");
        const existingData = sheet.getDataRange().getValues();
        const submitTimeStr = vendorData.submitTime.toISOString();
        for (let i = 1; i < existingData.length; i++) {
            if (existingData[i][0] === ticketNo && existingData[i][5]?.toISOString() === submitTimeStr) {
                Logger.log(`重複回應，跳過: ticketNo=${ticketNo}, submitTime=${submitTimeStr}`);
                return;
            }
        }

        // 寫入試算表
        sheet.appendRow([
            vendorData.ticketNo,
            vendorData.vendorName,
            vendorData.handlingDesc,
            vendorData.contactName,
            vendorData.contactPhone,
            vendorData.submitTime,
            vendorData.reply1,
            vendorData.reply2,
            vendorData.reply3,
            vendorData.reply4,
            vendorData.materialName,
            attachmentLink

        ]);
        Logger.log(`試算表寫入完成: ticketNo=${vendorData.ticketNo}`);

        // 記錄已處理的回應
        properties.setProperty(responseId, "processed");
        Logger.log(`標記回應為已處理: responseId=${responseId}`);

        // 寄發通知郵件
        sendNotificationEmail(ticketNo, vendorData, attachmentLink, notificationEmail, vendorEmail);

        // 刪除表單
        const formFile = DriveApp.getFileById(form.getId());
        formFile.setTrashed(true);
        Logger.log(`已刪除表單: ${formTitle}`);
        Logger.log(`處理完成: ticketNo=${ticketNo}`);

    } catch (error) {
        handleError("onFormSubmit", error, `表單標題: ${e.source?.getTitle()}`);
    }
}

/**
 * 處理檔案上傳
 */
function handleFileUpload(fileId, folderId) {
    if (!fileId) return "";
    try {
        const file = DriveApp.getFileById(fileId);
        const folder = DriveApp.getFolderById(folderId);
        const movedFile = file.moveTo(folder);
        return movedFile.getUrl();
    } catch (error) {
        Logger.log(`檔案上傳處理失敗: ${error.message}`);
        return "";
    }
}
 
function doGet(e) {
    try {
        const ticketNo = e.parameter.ticketNo;
        Logger.log(JSON.stringify(e));

        const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        Logger.log(spreadsheet);

        const sheet = spreadsheet.getSheetByName("廠商回應");
        const data = sheet.getDataRange().getValues();

        const response = data.filter((row) => String(row[0]).trim() === String(ticketNo).trim());
        if (response) {
            // 按 submitTime (row[5]) 降序排序，取得最新記錄
            const latestResponse = response.reduce((latest, current) => {
                const currentDate = new Date(current[5]);
                const latestDate = new Date(latest[5]);
                return currentDate > latestDate ? current : latest;
            });

            return ContentService.createTextOutput(
                JSON.stringify({
                    status: "success",
                    data: {
                        ticketNo: latestResponse[0],
                        vendorName: latestResponse[1],
                        handlingDesc: latestResponse[2],
                        contactName: latestResponse[3],
                        contactPhone: latestResponse[4],
                        submitTime: latestResponse[5],
                        reply1: latestResponse[6],
                        reply2: latestResponse[7],
                        reply3: latestResponse[8],
                        reply4: latestResponse[9],
                        attachmentLink: latestResponse[11]
                    },
                })
            ).setMimeType(ContentService.MimeType.JSON);
        } else {
            return ContentService.createTextOutput(
                JSON.stringify({
                    status: "error",
                    message: "未找到對應資料",
                })
            ).setMimeType(ContentService.MimeType.JSON);
        }
    } catch (error) {
        Logger.log(`doGet 錯誤: ${error.message}`);
        return ContentService.createTextOutput(
            JSON.stringify({
                status: "error",
                message: error.message,
            })
        ).setMimeType(ContentService.MimeType.JSON);
    }
}

/**
 * 寄送表單連結給廠商
 */
function sendFormEmail(vendorEmail, ticketNo, materialName, descript, prefilledUrl, vendorName, NOTIFICATION_EMAIL, NOTIFY_QA_EMAIL){
    try {
        const subject = `產品異常回覆單號 - ${ticketNo}`;
        const adminEmail = CONFIG.ADMIN_EMAIL;
        const recipient = encodeURIComponent(NOTIFICATION_EMAIL);
        const cc = encodeURIComponent(NOTIFY_QA_EMAIL);
        const mailtoLink = `mailto:${recipient}?cc=${cc}&subject=${subject}`;

        const htmlBody = `
            <html>
                <body>
                    <p>${vendorName} 您好，</p>
                    <p>請依以下說明填寫表單：</p>
                    <ul>
                        <li>產品異常回覆單號：${ticketNo}</li>
                        <li>物料名稱：${materialName}</li>
                        <li>異常描述：${descript}</li>
                    </ul>
                    <p><a href="${prefilledUrl}">請點我填寫表單</a></p>
                    <p><strong>附件提交方式：</strong><br>
                      → 寄送附件至 <a href="${mailtoLink}">${NOTIFICATION_EMAIL}</a>（主題包含單號 ${ticketNo}）；<br>
                      → 或上傳至雲端（Google Drive、Dropbox 等）並在表單「附件連結」欄位提供公開連結。<br>
                    - 若需 Google 帳戶，請<a href="https://accounts.google.com/signup">註冊</a>。
                    </p>
                    <p>如有問題，請聯繫管理員：${adminEmail}</p>
                    <p>感謝您的配合！</p>
                </body>
            </html>
        `;
        GmailApp.sendEmail(vendorEmail, subject, "", {
            htmlBody: htmlBody,
            name: "產品異常回覆單",
        });
        Logger.log(`表單連結已寄送至: ${vendorEmail}`);
    } catch (error) {
        handleError("sendFormEmail", error, `寄送對象: ${vendorEmail}`);
    }
}

/**
 * 寄送廠商回應通知郵件
 */
function sendNotificationEmail(ticketNo, vendorData, attachmentLink, notificationEmail, vendorEmail) {
    try {
        const subject = `產品異常回覆單 - 廠商回應 ${ticketNo}`;
        const htmlBody = `
            <html>
                <body>
                    <p>產品異常回覆單號: ${vendorData.ticketNo}</p>
                    <p>品名: ${vendorData.materialName}</p>
                    <p>廠商名稱: ${vendorData.vendorName}</p>
                    <p>異常處理說明: ${vendorData.handlingDesc}</p>
                    <p>聯絡人: ${vendorData.contactName}</p>
                    <p>聯絡電話: ${vendorData.contactPhone}</p>
                    ${
                        attachmentLink
                            ? `<p>附件連結: <a href="${attachmentLink}">${attachmentLink}</a></p>`
                            : "<p>附件：請檢查郵件（寄至 " + CONFIG.ADMIN_EMAIL + "）</p>"
                    }
                    <p><a href="https://bpm.company.com.tw/WebAgenda/index.do">請至系統頁面確認資料</a></p>


                     <p>提交時間: ${vendorData.submitTime}</p>
                </body>
            </html>
        `;

        const recipients = [notificationEmail || CONFIG.NOTIFICATION_EMAIL, vendorEmail || CONFIG.VENDOR_EMAIL]
            .filter(Boolean)
            .join(",");

        GmailApp.sendEmail(recipients, subject, "", {
            htmlBody: htmlBody,
            name: "產品異常回覆單",
        });
        Logger.log(`通知郵件寄送成功: ${recipients}`);

    } catch (error) {
        handleError("sendNotificationEmail", error, `產品異常回覆單號: ${ticketNo}`);
    }
}

/**
 * 每天下午 12 點執行，將試算表轉為 JSON 和 CSV
 */
function exportSheetsToJsonAndCsv() {
    try {
        const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        const sheetsToExport = ["表單紀錄", "廠商回應"];
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
        const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);

        sheetsToExport.forEach((sheetName) => {
            const sheet = spreadsheet.getSheetByName(sheetName);
            if (!sheet) {
                Logger.log(`工作表不存在: ${sheetName}`);
                return;
            }

            const data = sheet.getDataRange().getValues();
            if (data.length === 0) {
                Logger.log(`工作表無資料: ${sheetName}`);
                return;
            }

            const headers = data[0];
            const rows = data.slice(1);
            const jsonData = rows.map((row) => {
                let obj = {};
                headers.forEach((header, index) => {
                    obj[header] = row[index];
                });
                return obj;
            });
            const jsonString = JSON.stringify(jsonData, null, 2);
            const jsonFileName = `${sheetName}_${timestamp}.json`;
            folder.createFile(jsonFileName, jsonString, MimeType.PLAIN_TEXT);
            Logger.log(`JSON 檔案已儲存: ${jsonFileName}`);

            const csvData = [headers.join(",")]
                .concat(rows.map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(",")))
                .join("\n");
            const csvFileName = `${sheetName}_${timestamp}.csv`;
            folder.createFile(csvFileName, csvData, MimeType.CSV);
            Logger.log(`CSV 檔案已儲存: ${csvFileName}`);
        });

        // 寄送通知
        sendExportNotification(CONFIG.ADMIN_EMAIL, timestamp, folder.getUrl());
    } catch (error) {
        handleError("exportSheetsToJsonAndCsv", error);
    }
}

/**
 * 寄送匯出通知
 */
function sendExportNotification(recipient, timestamp, folderUrl) {
    try {
        const subject = `試算表匯出完成 - ${timestamp}`;
        const htmlBody = `
            <html>
                <body>
                    <p>您好，</p>
                    <p>試算表（表單紀錄、廠商回應）已於 ${timestamp} 成功匯出為 JSON 和 CSV 格式。</p>
                    <p>檔案已儲存至 Google 雲端硬碟：<br>
                       <a href="${folderUrl}">${folderUrl}</a></p>
                    <p>請下載檔案並匯入資料庫。</p>
                    <p>如有問題，請聯繫管理員：${CONFIG.ADMIN_EMAIL}</p>
                    <p>感謝您！</p>
                </body>
            </html>
        `;
        GmailApp.sendEmail(recipient, subject, "", {
            htmlBody: htmlBody,
            name: "產品異常回覆單",
        });
        Logger.log(`匯出通知已寄送至: ${recipient}`);
    } catch (error) {
        handleError("sendExportNotification", error, `時間戳: ${timestamp}`);
    }
}

/**
 * 每日 9 點將試算表下載為 Excel
 */
function exportSpreadsheetToDrive() {
    try {
        const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        const spreadsheetName = spreadsheet.getName();
        const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
        const fileName = `${spreadsheetName}_${date}.xlsx`;

        const url = `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/export?format=xlsx`;
        const token = ScriptApp.getOAuthToken();
        const response = UrlFetchApp.fetch(url, {
            headers: { Authorization: "Bearer " + token },
        });

        const blob = response.getBlob().setName(fileName);
        const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
        folder.createFile(blob);
        Logger.log(`Excel 檔案已儲存: ${fileName}`);
    } catch (error) {
        handleError("exportSpreadsheetToDrive", error);
    }
}

/**
 * 定期清理舊表單和附件（30 天前）
 */
function deleteOldFormsAndAttachments(days = 30) {
    try {
        const cutoffDate = new Date();
        cutoffDate.setDate(cutoffDate.getDate() - days);
        const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);

        // 清理舊表單
        const formFiles = DriveApp.getFilesByName("異常處理單 -");
        while (formFiles.hasNext()) {
            const file = formFiles.next();
            if (file.getLastUpdated() < cutoffDate) {
                file.setTrashed(true);
                Logger.log(`已刪除表單: ${file.getName()}`);
            }
        }

        // 清理舊附件（僅 JSON、CSV、Excel）
        const attachmentFiles = folder.getFiles();
        while (attachmentFiles.hasNext()) {
            const file = attachmentFiles.next();
            if (
                file.getLastUpdated() < cutoffDate &&
                (file.getName().endsWith(".json") ||
                    file.getName().endsWith(".csv") ||
                    file.getName().endsWith(".xlsx"))
            ) {
                file.setTrashed(true);
                Logger.log(`已刪除檔案: ${file.getName()}`);
            }
        }
    } catch (error) {
        handleError("deleteOldFormsAndAttachments", error);
    }
}

/**
 * 設定每日觸發器
 */
function createDailyTrigger() {
    // 刪除現有觸發器（除 onFormSubmit）
    ScriptApp.getProjectTriggers().forEach((trigger) => {
        if (trigger.getHandlerFunction() !== "onFormSubmit") {
            ScriptApp.deleteTrigger(trigger);
            Logger.log(`已刪除觸發器: ${trigger.getHandlerFunction()}`);
        }
    });

    // 創建新觸發器
    ScriptApp.newTrigger("exportSheetsToJsonAndCsv")
        .timeBased()
        .everyDays(1)
        .atHour(12)
        .nearMinute(15)
        .inTimezone("Asia/Taipei")
        .create();
    Logger.log("已設定 exportSheetsToJsonAndCsv 觸發器: 每日 12:15");

    ScriptApp.newTrigger("exportSpreadsheetToDrive")
        .timeBased()
        .everyDays(1)
        .atHour(9)
        .nearMinute(0)
        .inTimezone("Asia/Taipei")
        .create();
    Logger.log("已設定 exportSpreadsheetToDrive 觸發器: 每日 09:00");

    ScriptApp.newTrigger("deleteOldFormsAndAttachments")
        .timeBased()
        .everyDays(1)
        .atHour(2)
        .nearMinute(0)
        .inTimezone("Asia/Taipei")
        .create();
    Logger.log("已設定 deleteOldFormsAndAttachments 觸發器: 每日 02:00");
}
