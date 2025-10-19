function doGet() {
    // 回傳 index.html 作為首頁
    return HtmlService.createHtmlOutputFromFile("index");
}

/**
 * 取得目前登入使用者的 Email
 * @return {string} 使用者 Email
 */
function getUserEmail() {
    let email = Session.getActiveUser().getEmail();
    return email || "";
}

/**
 * 取得 Spreadsheet 實例，優先使用 getActiveSpreadsheet，失敗時從 Script Properties 的 SPREADSHEET_ID 回退
 * 若無法取得，會丟出錯誤
 */
function getSpreadsheet() {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        if (spreadsheet) return spreadsheet;
    } catch (err) {
        Logger.log(
            "getSpreadsheet: getActiveSpreadsheet error: %s",
            err && err.message
        );
    }
    try {
        const id =
            PropertiesService.getScriptProperties().getProperty(
                "SPREADSHEET_ID"
            );
        if (id) return SpreadsheetApp.openById(id);
    } catch (err) {
        Logger.log("getSpreadsheet: openById error: %s", err && err.message);
    }
    throw new Error(
        "無法存取試算表。請將此腳本綁定至對應試算表，或在 Script Properties 中設定 SPREADSHEET_ID，並確認已授權 spreadsheets 權限。"
    );
}

/**
 * 把 log 寫入名為 '管理日誌' 的工作表（若不存在會建立），欄位：時間、等級、訊息、額外 JSON
 */
function writeAdminLog(level, message, meta) {
    try {
        const ss = getSpreadsheet();
        let sh = ss.getSheetByName("管理日誌");
        if (!sh) sh = ss.insertSheet("管理日誌");
        const row = [
            new Date(),
            level,
            String(message || ""),
            meta ? JSON.stringify(meta) : "",
        ];
        sh.appendRow(row);
    } catch (err) {
        // 若寫入試算表失敗，退回到 Logger
        Logger.log("writeAdminLog failed: %s", err && err.message);
    }
}

/**
 * 取得前端登入設定，clientId 儲存在 Script Properties 的 GIS_CLIENT_ID
 * 回傳 { clientId: string, allowedDomain: string }
 */
function getAuthConfig() {
    const props = PropertiesService.getScriptProperties();
    const clientId = props.getProperty("GIS_CLIENT_ID") || "";
    let allowedDomain = "";

    try {
        const settings = getSettingsAsObject();
        const domain = settings["使用者信箱網域"];
        if (domain || domain === 0) {
            allowedDomain = String(domain).trim().toLowerCase();
        }
    } catch (err) {
        try {
            writeAdminLog("WARN", "getAuthConfig failed to read settings", {
                error: err && err.message,
            });
        } catch (logErr) {}
    }

    return {
        clientId: clientId,
        allowedDomain: allowedDomain,
    };
}

/**
 * 在 Script Properties 中寫入一個屬性（可透過 clasp run 呼叫）
 */
function setScriptProperty(key, value) {
    if (!key) throw new Error("key required");
    PropertiesService.getScriptProperties().setProperty(
        String(key),
        String(value)
    );
    Logger.log("setScriptProperty: %s set", key);
    try {
        writeAdminLog("INFO", "setScriptProperty: " + key, { value: "***" });
    } catch (err) {}
}

/**
 * 驗證 Google Identity Services 回傳的 ID token 並同步命題授權狀態
 * 回傳 getUserAccessInfo 的結果，附加 identityVerified 等欄位
 */
function verifyIdToken(idToken) {
    if (!idToken) throw new Error("缺少 ID token");

    const config = getAuthConfig();
    if (!config.clientId) {
        throw new Error("尚未於 Script Properties 設定 GIS_CLIENT_ID");
    }

    let response;
    try {
        response = UrlFetchApp.fetch(
            "https://oauth2.googleapis.com/tokeninfo?id_token=" +
                encodeURIComponent(idToken),
            { muteHttpExceptions: true }
        );
    } catch (err) {
        try {
            writeAdminLog("ERROR", "verifyIdToken fetch error", {
                error: err && err.message,
            });
        } catch (logErr) {}
        throw new Error("驗證 Google 登入失敗，請稍後再試");
    }

    if (response.getResponseCode() !== 200) {
        try {
            writeAdminLog("WARN", "verifyIdToken non-200", {
                status: response.getResponseCode(),
                body: response.getContentText(),
            });
        } catch (logErr) {}
        throw new Error("Google 登入資訊無效，請重新登入");
    }

    let payload;
    try {
        payload = JSON.parse(response.getContentText("utf-8"));
    } catch (err) {
        throw new Error("解析 Google 回傳資料時發生錯誤");
    }

    const issuer = String(payload.iss || "");
    const validIssuers = ["accounts.google.com", "https://accounts.google.com"];
    if (validIssuers.indexOf(issuer) < 0) {
        throw new Error("Google 登入來源不符");
    }

    const audience = String(payload.aud || "");
    if (audience !== config.clientId) {
        throw new Error("Google 登入 client_id 不符");
    }

    const email = String(payload.email || "");
    if (!email) {
        throw new Error("Google 登入未回傳 email");
    }

    const emailVerified =
        payload.email_verified === true || payload.email_verified === "true";
    if (!emailVerified) {
        throw new Error("Google 帳號尚未完成 email 驗證");
    }

    const nowSeconds = Math.floor(Date.now() / 1000);
    const expSeconds = Number(payload.exp || 0);
    if (expSeconds && expSeconds < nowSeconds) {
        throw new Error("Google 登入資訊已過期，請重新登入");
    }

    if (config.allowedDomain) {
        const normalizedAllowed = config.allowedDomain.toLowerCase();
        const emailDomain = email.split("@").pop().toLowerCase();
        const hostedDomain = String(payload.hd || "").toLowerCase();
        if (
            emailDomain !== normalizedAllowed &&
            hostedDomain !== normalizedAllowed
        ) {
            throw new Error("目前登入帳號的網域不在授權清單中");
        }
    }

    const accessInfo = getUserAccessInfo();
    if (!accessInfo.email) accessInfo.email = email;
    accessInfo.identityVerified = true;
    accessInfo.picture = payload.picture || "";
    accessInfo.fullName = payload.name || "";
    accessInfo.verifiedDomain = payload.hd || "";

    try {
        writeAdminLog("INFO", "verifyIdToken success", {
            email: email,
            authorized: accessInfo.authorized,
        });
    } catch (logErr) {}

    return accessInfo;
}

/**
 * 取得使用者 Email 並檢查是否出現在指定工作表的「命題教師Email」欄
 * 回傳物件 { email: string, authorized: boolean, domain: string }
 */
function getUserAccessInfo() {
    const rawEmail = Session.getActiveUser().getEmail() || "";
    const email = String(rawEmail);
    const emailLower = email.toLowerCase();

    const normalizeEmail = (input) => {
        if (!input && input !== 0) return "";
        const str = String(input).trim();
        const match = str.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
        if (match) return match[0].toLowerCase();
        const token = str.split(/[;|,\s()<>]+/)[0] || "";
        return token.toLowerCase();
    };

    const normalizeHeader = (h) => {
        if (h === null || h === undefined) return "";
        return String(h)
            .trim()
            .replace(/[\s\-_–—]+/g, "")
            .replace(/\uFEFF/g, "")
            .toLowerCase();
    };

    try {
        const ss = getSpreadsheet();
        if (!ss) return { email, authorized: false, domain: "" };

        const checkSheets = [
            "第1次定期考",
            "第2次定期考",
            "第3次定期考",
            "學期補考",
        ];

        // 一次性在記憶體讀取並整理每個工作表的資料，減少多次呼叫 Range
        const examTeachers = {}; // { sheetName: [ { email: normalizedEmail, name: rawName } ] }
        for (let i = 0; i < checkSheets.length; i++) {
            const sheetName = checkSheets[i];
            const sh = ss.getSheetByName(sheetName);
            if (!sh) continue;
            const data = sh.getDataRange().getValues();
            if (!data || data.length < 2) continue; // 沒有資料

            const headers = data[0].map(normalizeHeader);
            const emailColIndex = headers.findIndex((hh) =>
                /email|信箱|郵件/.test(hh)
            );
            if (emailColIndex < 0) continue;
            let nameColIndex = headers.findIndex((hh) => /姓名|名字/.test(hh));
            if (nameColIndex < 0) {
                if (emailColIndex - 1 >= 0) nameColIndex = emailColIndex - 1;
                else if (emailColIndex + 1 < headers.length)
                    nameColIndex = emailColIndex + 1;
            }

            examTeachers[sheetName] = [];
            for (let r = 1; r < data.length; r++) {
                const row = data[r];
                const emailCell = row[emailColIndex];
                if (!emailCell && emailCell !== 0) continue;
                const norm = normalizeEmail(emailCell);
                let nameVal = "";
                if (nameColIndex >= 0) {
                    const nameCell = row[nameColIndex];
                    if (nameCell || nameCell === 0)
                        nameVal = String(nameCell).trim();
                }
                examTeachers[sheetName].push({ email: norm, name: nameVal });
            }
        }

        // 在記憶體中尋找是否有匹配的命題教師
        const emailNorm = normalizeEmail(emailLower);
        let authorized = false;
        let teacherName = "";
        for (let i = 0; i < checkSheets.length; i++) {
            const sheetName = checkSheets[i];
            const list = examTeachers[sheetName];
            if (!list || !list.length) continue;
            for (let j = 0; j < list.length; j++) {
                const item = list[j];
                if (!item || !item.email) continue;
                if (item.email === emailNorm) {
                    authorized = true;
                    teacherName = item.name || "";
                    break;
                }
            }
            if (authorized) break;
        }

        const settings = getSettingsAsObject();
        const domain = settings["使用者信箱網域"] || "";

        const result = { email, authorized, domain, name: teacherName };

        try {
            writeAdminLog("DEBUG", "getUserAccessInfo batched read", {
                email: emailLower,
                authorized: authorized,
            });
        } catch (err) {}
        return result;
    } catch (errMain) {
        Logger.log(
            "getUserAccessInfo: error computing access info for %s: %s",
            emailLower,
            errMain && errMain.message
        );
        try {
            writeAdminLog("ERROR", "getUserAccessInfo exception", {
                email: emailLower,
                error: errMain && errMain.message,
            });
        } catch (err) {}
        return { email, authorized: false, domain: "" };
    }
}

/**
 * 讀取「設定」工作表的 A2:B 範圍，回傳物件 { A值: B值, ... }
 * 若找不到工作表或沒有資料，回傳空物件
 */
function getSettingsAsObject() {
    // try to obtain spreadsheet via helper if available
    let ss;
    try {
        // getUserAccessInfo defines getSpreadsheet in its scope; to be safe, re-implement fallback here
        try {
            ss = SpreadsheetApp.getActiveSpreadsheet();
        } catch (err) {
            Logger.log(
                "getSettingsAsObject: getActiveSpreadsheet error: %s",
                err && err.message
            );
            const id =
                PropertiesService.getScriptProperties().getProperty(
                    "SPREADSHEET_ID"
                );
            if (id) ss = SpreadsheetApp.openById(id);
        }
    } catch (err) {
        Logger.log(
            "getSettingsAsObject: cannot access spreadsheet: %s",
            err && err.message
        );
        try {
            writeAdminLog(
                "ERROR",
                "getSettingsAsObject cannot access spreadsheet",
                { error: err && err.message }
            );
        } catch (err2) {}
        return {};
    }
    if (!ss) return {};

    const sh = ss.getSheetByName("設定");
    if (!sh) return {};

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return {};

    const rows = sh.getRange(2, 1, lastRow - 1, 2).getValues();
    try {
        writeAdminLog("DEBUG", "getSettingsAsObject read rows", {
            rows: rows.length,
        });
    } catch (err) {}
    return rows.reduce((acc, rowArr) => {
        const rawKey = rowArr[0];
        let valueCell = rowArr[1];
        const key =
            rawKey === null || rawKey === undefined
                ? ""
                : String(rawKey).trim();
        if (key === "") return acc;

        // normalize strings
        if (typeof valueCell === "string") {
            const valueStr = valueCell.trim();
            // boolean
            const lv = valueStr.toLowerCase();
            if (lv === "true") valueCell = true;
            else if (lv === "false") valueCell = false;
            // number
            else if (/^-?\d+(?:\.\d+)?$/.test(valueStr))
                valueCell = Number(valueStr);
            else valueCell = valueStr;
        }

        acc[key] = valueCell;
        try {
            writeAdminLog("DEBUG", "setting parsed", {
                key: key,
                value: valueCell,
            });
        } catch (err) {}
        return acc;
    }, {});
}

function test() {
    Logger.log(getSettingsAsObject());
    Logger.log(getUserAccessInfo());
}
