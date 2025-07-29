# Google Apps Script — 產品異常回覆自動化系統  

以 Google Apps Script 為基礎，整合 Google Forms、Sheets、Email 與 Web API，專為 BPM 系統建置的自動化回覆平台。  

---

##  開發目的 | Purpose

- 減少人工建立 Google 表單與通知流程  
- 提供標準 API，建立及查詢異常資料
- 使用 GmailApp.sendEmai 寄送異常資料及廠商回覆通知  
  

##  主要功能 | Key Features

-  **POST 建立異常表單**  
  BPM 傳送異常資料 → 自動建立 Google 表單 → 通知廠商  

-  **GET 查詢異常回覆**  
  BPM 依據單號查詢 → 取得廠商填寫內容 → 更新系統狀態  

-  **資料記錄與備份**  
  表單資料即時寫入 Google Sheets，並定期匯出 Excel / JSON / CSV  

-  **自動清理與排程任務**  
  排程清除 30 天以上舊資料與附件  


---

##  API 快速總覽 | API Summary

| 功能 | Method | 說明 |
|------|--------|------|
| 建立表單 | `POST /doPost` | 建立 Google 表單並寄送廠商 | 
| 查詢回覆 | `GET /doGet?ticketNo=...` | 查詢指定單號的回覆資料 | 

 

## 系統架構與作業流程

### 1. POST /doPost 通知廠商流程
```
BPM System → Google Apps Script API → Google Forms → Google Sheet → 郵件通知廠商
```

**流程說明**：
1. BPM 系統產品異常，透過 POST API 傳送異常資料
2. Google Apps Script 接收資料，自動創建客製化表單，內崁圖片
3. 系統自動寄送表單連結給廠商
4. 廠商填寫異常處理回覆（原因分析、矯正措施、預防措施等）
5. 表單提交後自動記錄至試算表
6. 系統自動通知相關人員（採購、品保）

### 2. GET 確認資料流程
```
BPM System → Google Apps Script API → 廠商回應資料 → Google Sheet →  郵件通知人員確認 → BPM System
```

**流程說明**：
1. BPM 系統確認廠商回應狀態時，透過 GET API 查詢
2. Google Apps Script 從試算表中搜尋對應單號的廠商回應資料
3. 系統回傳最新的廠商回應內容（包含處理說明、聯絡人、各項措施等）

## 配置設定

### CONFIG 常數說明

```javascript
const CONFIG = {
    DRIVE_FOLDER_ID: "1QCFpQb9XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", // 雲端硬碟資料夾
    SPREADSHEET_ID: "1CG9xXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX", // 試算表 ID
    ADMIN_EMAIL: "notify.admin@gmail.com", // 管理員郵件
    VENDOR_EMAIL: "notify.vendor@gmail.com.tw", // 預設廠商郵件
    NOTIFICATION_EMAIL: "notify.bpm@gmail.com", // 通知郵件
    NOTIFY_QA_EMAIL: "notify.QA@gmail.com" // 品保通知郵件
};
```

## API 介接規格

### 1. doPost() - 創建表單 API  

**端點**: `https://script.google.com/macros/s/{SCRIPT_ID}/exec`  

**方法**: POST  

**Content-Type**: application/json  

**用途**: 供 BPM 系統傳送異常資料，自動創建表單並通知廠商  


#### 使用場景
- BPM 系統 - 產品異常需要廠商回覆
- 自動創建客製化的異常回覆表單
- 寄送表單連結給指定廠商
- 記錄異常處理資訊

#### 請求參數

| 參數名稱 | 類型 | 必填 | 說明 |
|---------|------|------|------|
| ticketNo | string | ✓ | 異常處理單號 |
| materialName | string | ✓ | 物料名稱 |
| descript | string | ✓ | 異常描述 |
| vendorEmail | string | ✓ | 廠商郵件 |
| aplyDate | string | ✗ | 申請日期 |
| abnItem | string | ✗ | 異常項目 |
| imageUrl | string | ✗ | 異常圖片路徑 (支援多張，以 ; 分隔) |
| notificationEmail | string | ✗ | 通知郵件 |
| notifyQaEmail | string | ✗ | 品保通知郵件 |
| vendorName | string | ✗ | 廠商名稱 |

#### 請求範例

```json
{
    "ticketNo": "ART03161724310232159",
    "materialName": "拉蝦",
    "descript": "產品品質異常",
    "vendorEmail": "vendor@example.com",
    "aplyDate": "2025-07-03",
    "abnItem": "產品異常",
    "imageUrl": "../myproj/ART03161724310232159/image1.jpg;../myproj/ART03161724310232159/image2.jpg",
    "notificationEmail": "qa@company.com",
    "notifyQaEmail": "purchase@company.com",
    "vendorName": "ABC供應商"
}
```

#### 回應格式

**成功回應**:
```json
{
    "status": "success",
    "formUrl": "https://docs.google.com/forms/d/e/1FAIpQLSd.../viewform"
}
```

**錯誤回應**:
```json
{
    "status": "error",
    "message": "缺少必要欄位:單號、料名 或 異常描述、廠商email"
}
```

### 2. doGet() - 查詢廠商回應 API

**端點**: `https://script.google.com/macros/s/{SCRIPT_ID}/exec?ticketNo={TICKET_NO}`  

**方法**: GET  

**用途**: 供 BPM 系統查詢廠商回應資料，確認異常處理狀態  

#### 使用場景
- BPM 系統確認廠商回覆異常處理
- 查詢廠商回應的具體內容（原因分析、矯正措施等）
- 更新 BPM 系統產品異常 責任單位回覆
- 進行後續品保追蹤作業

#### 請求參數

| 參數名稱 | 類型 | 必填 | 說明 |
|---------|------|------|------|
| ticketNo | string | ✓ | 異常處理單號 |

#### 查詢邏輯
- 系統會搜尋「廠商回應」工作表中對應的單號
- 如有多筆回應，會回傳最新的一筆（依提交時間排序）
- 回傳完整的廠商回應資料供 BPM 系統處理

#### 回應格式

**成功回應**:
```json
{
    "status": "success",
    "data": {
        "ticketNo": "ART03161724310232159",
        "vendorName": "ABC供應商",
        "handlingDesc": "已緊急更換供應來源",
        "contactName": "張三",
        "contactPhone": "0912345678",
        "submitTime": "2025-07-08T10:30:00.000Z",
        "reply1": "立即停止出貨，回收問題批次",
        "reply2": "原料供應商品質管控不當導致",
        "reply3": "重新制定供應商審核標準",
        "reply4": "建立每批檢驗制度",
        "attachmentLink": "https://drive.google.com/file/d/xxx/view"
    }
}
```

**未找到資料**:
```json
{
    "status": "error",
    "message": "未找到對應資料"
}
```

**系統錯誤**:
```json
{
    "status": "error",
    "message": "系統處理錯誤訊息"
}
```

## 試算表結構

### 表單紀錄工作表

| 欄位 | 說明 |
|------|------|
| A | 單號 (ticketNo) |
| B | 申請日期 (aplyDate) |
| C | 異常項目 (abnItem) |
| D | 物料名稱 (materialName) |
| E | 異常描述 (descript) |
| F | 廠商名稱 (vendorName) |
| G | 圖片路徑 (imageUrl) |
| H | 表單連結 |
| I | 建立時間 |
| J | 廠商郵件 |
| K | 通知郵件 |
| L | 品保郵件 |

### 廠商回應工作表

| 欄位 | 說明 |
|------|------|
| A | 單號 |
| B | 廠商名稱 |
| C | 異常處理說明 |
| D | 聯絡人姓名 |
| E | 聯絡電話 |
| F | 提交時間 |
| G | 應急措施 |
| H | 原因分析 |
| I | 矯正措施 |
| J | 預防措施 |
| K | 物料名稱 |
| L | 附件連結 |

## 自動化排程

### 定期任務設定

1. **每日 02:00** - `deleteOldFormsAndAttachments()`
   - 清理 30 天前的舊表單和附件
   - 避免 Google Drive 空間不足

2. **每日 09:00** - `exportSpreadsheetToDrive()`
   - 匯出試算表為 Excel 格式
   - 儲存至指定 Google Drive 資料夾

3. **每日 12:15** - `exportSheetsToJsonAndCsv()`
   - 將試算表匯出為 JSON 和 CSV 格式
   - 寄送通知郵件給管理員

### 觸發器設定

執行以下函數來設定所有定期任務：

```javascript
createDailyTrigger();
```

## 維運指南

### 1. 部署和設定

1. **開啟 Google Apps Script**
   - 前往 [script.google.com](https://script.google.com)
   - 創建新專案，貼上程式碼

2. **設定配置**
   - 修改 `CONFIG` 常數中的各項設定
   - 確認 Google Drive 資料夾和試算表 ID

3. **部署 Web 應用程式**
   - 點擊「部署」→「新增部署」
   - 選擇「Web 應用程式」
   - 執行身分：我
   - 存取權限：任何人 (用於接收 BPM 系統請求)

4. **設定觸發器**
   - 執行 `createDailyTrigger()` 函數
   - 確認觸發器設定成功

### 2. 監控和維護

#### 日誌監控

1. **執行紀錄**
   - 前往 Apps Script 編輯器
   - 點擊「執行」→「執行紀錄」
   - 監控 API 調用和錯誤

2. **錯誤通知**
   - 系統會自動寄送錯誤通知到管理員信箱
   - 檢查 `handleError()` 函數的通知

#### 常見問題處理

1. **表單創建失敗**
   ```javascript
   // 檢查錯誤原因
   Logger.log("檢查 ticketNo, materialName, descript 是否完整");
   ```

2. **郵件寄送失敗**
   - 檢查郵件配額 (個人帳戶每天 100 封)
   - 確認收件人郵件格式正確

3. **觸發器失效**
   ```javascript
   // 重新建立觸發器
   createDailyTrigger();
   ```

### 3. 資料備份

1. **自動備份**
   - 系統每日自動匯出 JSON、CSV、Excel 格式
   - 檔案儲存於指定的 Google Drive 資料夾

2. **手動備份**
   ```javascript
   // 手動執行匯出
   exportSheetsToJsonAndCsv();
   exportSpreadsheetToDrive();
   ```

### 4. 容量管理

1. **清理策略**
   - 30 天自動清理舊表單和附件
   - 手動清理可調整天數：`deleteOldFormsAndAttachments(7)`

2. **配額限制**
   - Gmail 寄送：100 封/天 (個人帳戶)
   - Google Drive：15 GB 免費空間
   - Apps Script：6 分鐘/執行，30 次/分鐘

## 安全考量

1. **權限控制**
   - Web 應用程式設為「任何人」存取 
   - 建議使用 API 金鑰驗證 (需自行實作)

2. **資料保護**
   - 敏感資料存於 Google 雲端
   - 定期備份重要資料

3. **錯誤處理**
   - 所有錯誤都會記錄並通知管理員
   - 防止重複表單提交

## 開發和擴展

### 自定義修改

1. **新增表單欄位**
   ```javascript
   // 在 createForm() 函數中添加
   form.addTextItem().setTitle("新欄位").setRequired(true);
   ```

2. **修改通知郵件**
   - 編輯 `sendFormEmail()` 和 `sendNotificationEmail()` 函數
   - 調整郵件模板和收件人

3. **數據處理**
   - 修改 `onFormSubmit()` 函數處理新欄位
   - 調整試算表欄位對應

### 測試建議

1. **API 測試**
   - 使用 Postman 或 curl 測試 API 端點
   - 驗證各種參數組合

2. **表單測試**
   - 測試表單填寫和提交流程
   - 確認郵件通知正常

3. **排程測試**
   - 手動執行定期任務
   - 確認備份和清理功能

## 更新紀錄
- **2025-07-02**: 移除 GmailApp.getRemainingDailyQuota (個人帳戶不支援)

---

*最後更新：2025-07-15*
