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
        // 嘗試擷取符合 RFC 的 email 部分
        const match = str.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
        if (match) return match[0].toLowerCase();
        // 若無符合 email 的部分，備援：以常見分隔符切割並取第一段
        const token = str.split(/[;|,\s()<>]+/)[0] || "";
        return token.toLowerCase();
    };

    // (已移除 CacheService 使用：改為直接查詢試算表並回傳最新資訊)

    try {
        const ss = getSpreadsheet();
        if (!ss) return { email, authorized: false, domain: "" };

        const checkSheets = [
            "第1次定期考",
            "第2次定期考",
            "第3次定期考",
            "學期補考",
        ];
        const emailNorm = normalizeEmail(emailLower);
        // 嘗試取得對應的命題教師姓名（若 email 匹配）
        let authorized = false;
        let teacherName = "";
        for (let i = 0; i < checkSheets.length; i++) {
            const sheetName = checkSheets[i];
            const sh = ss.getSheetByName(sheetName);
            if (!sh) continue;
            const lastCol = sh.getLastColumn();
            if (lastCol < 1) continue;
            const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
            // 正規化標題：去掉空白、破折號、底線，轉小寫，方便容錯匹配
            const normalizeHeader = (h) => {
                if (h === null || h === undefined) return "";
                return String(h)
                    .trim()
                    .replace(/[\s\-_–—]+/g, "")
                    .replace(/\uFEFF/g, "")
                    .toLowerCase();
            };
            const normHeaders = headers.map(normalizeHeader);
            // 容錯：email 欄可能標題包含 email / 信箱 / 郵件 等字詞
            const emailColIndex = normHeaders.findIndex((hh) =>
                /email|信箱|郵件/.test(hh)
            );
            if (emailColIndex < 0) continue;
            // 容錯：姓名欄可能以「姓名」「名字」等命名
            let nameColIndex = normHeaders.findIndex((hh) =>
                /姓名|名字/.test(hh)
            );
            // 若找不到明確的姓名欄，嘗試使用 email 欄左右相鄰欄作為候補（常見情況）
            if (nameColIndex < 0) {
                if (emailColIndex - 1 >= 0) nameColIndex = emailColIndex - 1;
                else if (emailColIndex + 1 < normHeaders.length)
                    nameColIndex = emailColIndex + 1;
            }
            const lastRow = sh.getLastRow();
            if (lastRow <= 1) continue;
            const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
            for (let r = 0; r < rows.length; r++) {
                const row = rows[r];
                const emailCell = row[emailColIndex];
                if (!emailCell && emailCell !== 0) continue;
                if (normalizeEmail(emailCell) === emailNorm) {
                    authorized = true;
                    if (nameColIndex >= 0) {
                        const nameCell = row[nameColIndex];
                        if (nameCell || nameCell === 0)
                            teacherName = String(nameCell).trim();
                    }
                    break;
                }
            }
            if (authorized) break;
        }

        const settings = getSettingsAsObject();
        const domain = settings["使用者信箱網域"] || "";

        const result = { email, authorized, domain, name: teacherName };

        // 不使用 CacheService：直接回傳最新查詢結果
        try {
            writeAdminLog("DEBUG", "getUserAccessInfo computed", {
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
        // 發生錯誤時回傳安全的預設值
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
