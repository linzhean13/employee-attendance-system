# 🏢 員工出勤管理系統 (Employee Attendance System)

> 基於 C# Windows Forms 開發的輕量級企業出勤打卡與人事管理系統，實現「零配置」的自動化資料庫建置。

## 📝 專案簡介
本專案為一套獨立運作的桌面端出勤系統，分為「員工打卡前台」與「管理員後台」。系統無需額外架設伺服器，利用 ADO.NET 技術串接 Microsoft Access 資料庫，並具備**自動建置資料庫與防呆驗證機制**，確保差勤數據的準確性與系統的易用性。

## 🛠️ 技術標籤 (Tech Stack)
* **開發語言:** C# (.NET Framework)
* **應用程式框架:** Windows Forms (WinForms)
* **資料庫:** Microsoft Access (`.mdb`)
* **資料存取技術:** ADO.NET (OLE DB), ADOX (動態建立資料庫)
* **UI 開發方式:** Programmatic UI (純程式碼動態生成介面元件)

## ✨ 核心功能與技術亮點

### 🌟 系統特色
* **零配置啟動 (Zero-Configuration)：** 系統具備環境自檢能力。首次啟動時若偵測無資料庫，將透過 `ADOX` 自動建立 `.mdb` 檔案、構建資料表架構，並注入測試員工資料，實現真正的「開箱即用」。
* **純程式碼建構 UI：** 捨棄傳統 Designer 拖曳，全系統介面（包含動態座標、大小、事件綁定）皆透過 C# 物件導向程式碼即時生成，展現對 UI 框架底層運作的掌握度。

### 👤 員工前台 (Frontend)
* **狀態防呆機制：** 系統會動態比對當日最後一筆紀錄狀態，嚴格防堵「未下班重複打上班卡」或「未上班直接打下班卡」等異常操作。
* **即時出勤看板：** 透過 `DataGridView` 綁定資料表，打卡後即時刷新全體員工的當日出勤動態與時間戳記。

### 👨‍💼 管理後台 (Admin Dashboard)
* **身分驗證：** 具備簡易的管理員登入鎖定機制（內建測試密碼）。
* **多條件交叉查詢：** 管理員可透過「員工姓名」與「自訂日期區間 (Date Range)」進行組合查詢，快速調閱特定區間的出勤報表。
* **人事管理系統：** 提供員工的新增與刪除功能，並在新增前執行 `SQL COUNT(*)` 查重，防止建立同名員工資料。

## 📂 專案架構
```text
EmployeeAttendance/
│
├── EmployeeAttendance.slnx      # Visual Studio 方案檔 (請點擊此檔開啟專案)
│
└── EmployeeAttendance/         # 主程式專案目錄
    ├── Form1.cs                # 系統入口與員工打卡前台邏輯
    ├── Form2.cs                # 管理員後台與報表查詢邏輯
    ├── Program.cs              # 應用程式啟動點
    └── ... (其他 .NET 方案設定檔)
```
> 💡 *註：`AttendanceSystem.mdb` 資料庫檔案會在系統首次編譯執行時，自動生成於 `bin/Debug` 目錄下。*

## 🚀 如何在本地端執行 (How to Run)

1. **環境要求：** 請確保電腦已安裝 [Visual Studio](https://visualstudio.microsoft.com/zh-hant/) (支援 .NET 桌面開發工作負載)。
2. **開啟專案：** 下載本專案後，雙擊 `EmployeeAttendance.slnx` 檔案以 Visual Studio 開啟。
3. **編譯執行：** 點擊上方工具列的「開始」或按下 `F5` 啟動應用程式。
4. **測試操作：**
   * **員工打卡：** 預設已建立 3 名測試員工，可直接於下拉選單選擇並測試上下班打卡。
   * **進入後台：** 點選上方「管理員」選項，輸入測試密碼 `1234` 即可進入後台管理介面。

## 📸 系統畫面展示 (Screenshots)

<table width="100%">
  <tr>
    <td width="50%" valign="top">
      <b>1. 員工打卡前台 (Form1)</b><br>
      員工可選擇姓名進行上下班打卡，下方清單即時顯示最新出勤狀態。系統內建邏輯防止重複打卡錯誤。
      <br><br>
      <img src="請貼上你的前台截圖網址" width="100%">
    </td>
    <td width="50%" valign="top">
      <b>2. 管理員後台 (Form2)</b><br>
      提供指定日期區間與特定員工的差勤過濾查詢，並支援右側的人事資料管理（新增/刪除員工）。
      <br><br>
      <img src="請貼上你的後台截圖網址" width="100%">
    </td>
  </tr>
</table>
