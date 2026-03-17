using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;

namespace EmployeeAttendance
{
    public partial class Form2 : Form
    {
        // --- 介面元件 ---
        private Label lblFilter;
        private ComboBox cmbEmployeeFilter;

        private Label lblDateRange;
        private DateTimePicker dtpStart;
        private Label lblTo;
        private DateTimePicker dtpEnd;
        private Button btnSearch;

        private GroupBox grpManage;
        private Label lblAddHint;
        private TextBox txbNewName;
        private Button btnAdd;
        private Button btnDelete;

        private DataGridView dgvData;
        private Button btnBack;

        // 資料庫設定
        string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=AttendanceSystem.mdb;";

        public Form2()
        {
            InitializeComponent();
            SetupUI();
            SetDefaultDateRange();
            LoadFilterList();
            ShowRecords();
        }

        private void SetDefaultDateRange()
        {
            DateTime now = DateTime.Now;
            DateTime firstDay = new DateTime(now.Year, now.Month, 1);
            DateTime lastDay = firstDay.AddMonths(1).AddDays(-1);
            dtpStart.Value = firstDay;
            dtpEnd.Value = lastDay;
        }

        private void SetupUI()
        {
            // 視窗設定
            this.Size = new Size(1000, 750);
            this.Text = "員工出勤系統 - 管理後台";
            this.StartPosition = FormStartPosition.CenterScreen;

            // --- 篩選區 ---
            lblFilter = new Label() { Text = "員工篩選：", Location = new Point(20, 25), AutoSize = true, Font = new Font("微軟正黑體", 10) };
            this.Controls.Add(lblFilter);

            cmbEmployeeFilter = new ComboBox() { Location = new Point(100, 22), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbEmployeeFilter.SelectedIndexChanged += (s, e) => ShowRecords();
            this.Controls.Add(cmbEmployeeFilter);

            lblDateRange = new Label() { Text = "日期區間：", Location = new Point(20, 65), AutoSize = true, Font = new Font("微軟正黑體", 10) };
            this.Controls.Add(lblDateRange);

            dtpStart = new DateTimePicker() { Location = new Point(100, 62), Width = 120, Format = DateTimePickerFormat.Short };
            this.Controls.Add(dtpStart);

            lblTo = new Label() { Text = "~", Location = new Point(225, 65), AutoSize = true };
            this.Controls.Add(lblTo);

            dtpEnd = new DateTimePicker() { Location = new Point(245, 62), Width = 120, Format = DateTimePickerFormat.Short };
            this.Controls.Add(dtpEnd);

            btnSearch = new Button() { Text = "查詢區間", Location = new Point(380, 60), Size = new Size(80, 27), BackColor = Color.LightSkyBlue };
            btnSearch.Click += (s, e) => ShowRecords();
            this.Controls.Add(btnSearch);

            // --- 管理區 ---
            grpManage = new GroupBox() { Text = "人員管理", Location = new Point(520, 10), Size = new Size(420, 80) };
            this.Controls.Add(grpManage);

            lblAddHint = new Label() { Text = "姓名：", Location = new Point(15, 30), AutoSize = true };
            grpManage.Controls.Add(lblAddHint);

            txbNewName = new TextBox() { Location = new Point(60, 27), Width = 100 };
            grpManage.Controls.Add(txbNewName);

            btnAdd = new Button() { Text = "新增", Location = new Point(170, 25), Size = new Size(60, 25), BackColor = Color.LightGreen };
            btnAdd.Click += new EventHandler(btnAdd_Click);
            grpManage.Controls.Add(btnAdd);

            btnDelete = new Button() { Text = "刪除(左側選取者)", Location = new Point(240, 25), Size = new Size(140, 25), ForeColor = Color.Red };
            btnDelete.Click += new EventHandler(btnDelete_Click);
            grpManage.Controls.Add(btnDelete);

            // --- 資料表格 ---
            dgvData = new DataGridView();
            dgvData.Location = new Point(20, 110);
            dgvData.Size = new Size(940, 540);
            dgvData.ReadOnly = true;
            dgvData.AllowUserToAddRows = false;
            dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvData.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.Controls.Add(dgvData);

            // --- 登出按鈕 ---
            btnBack = new Button() { Text = "登出", Size = new Size(100, 40) };
            btnBack.Location = new Point(this.ClientSize.Width - 140, this.ClientSize.Height - 60);
            btnBack.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnBack.Click += new EventHandler(btnBack_Click);
            this.Controls.Add(btnBack);
        }

        private void LoadFilterList()
        {
            try
            {
                using (OleDbConnection conn = new OleDbConnection(connStr))
                {
                    conn.Open();
                    string sql = "SELECT EmpID, EmpName FROM Employee ORDER BY EmpID";
                    OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    DataRow allRow = dt.NewRow();
                    allRow["EmpID"] = -1;
                    allRow["EmpName"] = "(顯示所有員工)";
                    dt.Rows.InsertAt(allRow, 0);

                    cmbEmployeeFilter.DataSource = dt;
                    cmbEmployeeFilter.DisplayMember = "EmpName";
                    cmbEmployeeFilter.ValueMember = "EmpID";
                }
            }
            catch (Exception ex) { MessageBox.Show("讀取名單失敗: " + ex.Message); }
        }

        private void ShowRecords()
        {
            if (cmbEmployeeFilter.SelectedValue == null) return;

            try
            {
                int selectedId = -1;
                int.TryParse(cmbEmployeeFilter.SelectedValue.ToString(), out selectedId);

                string strStart = dtpStart.Value.ToString("yyyy/MM/dd");
                string strEnd = dtpEnd.Value.ToString("yyyy/MM/dd");

                using (OleDbConnection conn = new OleDbConnection(connStr))
                {
                    conn.Open();
                    string sql = "SELECT Employee.EmpName AS 姓名, " +
                                 "DailyAttendance.WorkDate AS 日期, " +
                                 "DailyAttendance.StartTime AS 上班時間, " +
                                 "DailyAttendance.EndTime AS 下班時間, " +
                                 "DailyAttendance.Status AS 狀態 " +
                                 "FROM DailyAttendance, Employee " +
                                 "WHERE DailyAttendance.EmpID = Employee.EmpID ";

                    sql += $" AND DailyAttendance.WorkDate >= '{strStart}' AND DailyAttendance.WorkDate <= '{strEnd}'";

                    if (selectedId != -1)
                    {
                        sql += $" AND Employee.EmpID = {selectedId}";
                    }

                    sql += " ORDER BY DailyAttendance.ID DESC";

                    OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dgvData.DataSource = dt;
                }
            }
            catch (Exception ex) { MessageBox.Show("查詢失敗: " + ex.Message); }
        }

        // 檢查員工是否存在
        private bool CheckIfEmployeeExists(string name)
        {
            bool exists = false;
            try
            {
                using (OleDbConnection conn = new OleDbConnection(connStr))
                {
                    conn.Open();
                    // 使用 COUNT(*) 來計算有幾筆名字相同的資料
                    string sql = $"SELECT COUNT(*) FROM Employee WHERE EmpName = '{name}'";
                    OleDbCommand cmd = new OleDbCommand(sql, conn);
                    int count = (int)cmd.ExecuteScalar(); // 取得查詢結果的第一個數值
                    if (count > 0)
                    {
                        exists = true;
                    }
                }
            }
            catch { }
            return exists;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            string newName = txbNewName.Text.Trim(); // 去除前後空白
            if (string.IsNullOrEmpty(newName)) { MessageBox.Show("請輸入姓名"); return; }

            // 1. 先檢查是否重複
            if (CheckIfEmployeeExists(newName))
            {
                MessageBox.Show($"錯誤：員工 [{newName}] 已經存在，無法重複新增！", "重複警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // 直接結束，不執行下面的 Insert
            }

            // 2. 沒重複才執行新增
            if (TryExecuteSql($"INSERT INTO Employee (EmpName) VALUES ('{newName}')"))
            {
                MessageBox.Show($"新增 {newName} 成功！");
                txbNewName.Text = "";
                LoadFilterList();
                ShowRecords();
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (cmbEmployeeFilter.SelectedValue == null) return;
            int selectedId = (int)cmbEmployeeFilter.SelectedValue;

            if (selectedId == -1)
            {
                MessageBox.Show("請先在左側選單選擇一位具體的員工，才能進行刪除。");
                return;
            }

            string empName = cmbEmployeeFilter.Text;
            if (MessageBox.Show($"確定要刪除 [{empName}] 嗎？", "刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                if (TryExecuteSql($"DELETE FROM Employee WHERE EmpID = {selectedId}"))
                {
                    MessageBox.Show("刪除成功");
                    LoadFilterList();
                    ShowRecords();
                }
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public bool TryExecuteSql(string sql)
        {
            try
            {
                using (OleDbConnection conn = new OleDbConnection(connStr))
                {
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand(sql, conn);
                    cmd.ExecuteNonQuery();
                    return true;
                }
            }
            catch (Exception ex) { MessageBox.Show("資料庫錯誤：" + ex.Message); return false; }
        }
    }
}