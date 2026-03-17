using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using ADOX;

namespace EmployeeAttendance
{
    public partial class Form1 : Form
    {
        // --- 介面元件 ---
        private Label lblTitle;
        private RadioButton rbEmployee;
        private RadioButton rbManager;
        private Label lblEmpName;
        private ComboBox cmbEmployee;
        private Button btnClockIn;
        private Button btnClockOut;
        private Label lblPwd;
        private TextBox txbPassword;
        private Button btnLogin;
        private DataGridView dataGridView1;

        // 資料庫設定
        string dbname = "AttendanceSystem.mdb";
        string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=AttendanceSystem.mdb;";

        OleDbConnectionStringBuilder connstring = null;
        OleDbConnection conn;

        public Form1()
        {
            InitializeComponent();
            CheckAndCreateDatabase();
            SetupUI();
            SetupDatabase();
            LoadEmployeeList();

            // 預設顯示員工模式
            rbEmployee.Checked = true;
            ShowLog();
        }

        // --- 資料庫初始化 ---
        private void CheckAndCreateDatabase()
        {
            if (!File.Exists(dbname))
            {
                try
                {
                    Catalog cat = new Catalog();
                    cat.Create(connStr);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cat.ActiveConnection);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cat);

                    using (OleDbConnection conn = new OleDbConnection(connStr))
                    {
                        conn.Open();
                        OleDbCommand cmd = new OleDbCommand();
                        cmd.Connection = conn;
                        cmd.CommandText = "CREATE TABLE Employee ([EmpID] AUTOINCREMENT PRIMARY KEY, [EmpName] VARCHAR(50))";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "CREATE TABLE DailyAttendance ([ID] AUTOINCREMENT PRIMARY KEY, [EmpID] INT, [WorkDate] VARCHAR(20), [StartTime] VARCHAR(20), [EndTime] VARCHAR(20), [Status] VARCHAR(20))";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "INSERT INTO Employee (EmpName) VALUES ('人名001')";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "INSERT INTO Employee (EmpName) VALUES ('人名002')";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "INSERT INTO Employee (EmpName) VALUES ('人名003')";
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex) { MessageBox.Show("資料庫建立失敗：" + ex.Message); }
            }
        }

        // --- 畫面佈局 ---
        private void SetupUI()
        {
            this.Size = new Size(800, 600);
            this.Text = "員工出勤系統 - 打卡前台";

            lblTitle = new Label() { Text = "請選擇身份：", Location = new Point(20, 20), AutoSize = true, Font = new Font("微軟正黑體", 12) };
            this.Controls.Add(lblTitle);

            rbEmployee = new RadioButton() { Text = "一般員工", Location = new Point(130, 20), AutoSize = true };
            rbEmployee.CheckedChanged += new EventHandler(rb_CheckedChanged);
            this.Controls.Add(rbEmployee);

            rbManager = new RadioButton() { Text = "管理員", Location = new Point(220, 20), AutoSize = true };
            rbManager.CheckedChanged += new EventHandler(rb_CheckedChanged);
            this.Controls.Add(rbManager);

            lblEmpName = new Label() { Text = "請選擇姓名：", Location = new Point(20, 60), AutoSize = true };
            this.Controls.Add(lblEmpName);

            cmbEmployee = new ComboBox() { Location = new Point(110, 57), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            this.Controls.Add(cmbEmployee);

            btnClockIn = new Button() { Text = "上班", Location = new Point(270, 55), Size = new Size(80, 30), BackColor = Color.LightGreen };
            btnClockIn.Click += new EventHandler(btnClockIn_Click);
            this.Controls.Add(btnClockIn);

            btnClockOut = new Button() { Text = "下班", Location = new Point(360, 55), Size = new Size(80, 30), BackColor = Color.LightSalmon };
            btnClockOut.Click += new EventHandler(btnClockOut_Click);
            this.Controls.Add(btnClockOut);

            lblPwd = new Label() { Text = "管理員密碼：", Location = new Point(20, 60), AutoSize = true, Visible = false };
            this.Controls.Add(lblPwd);

            txbPassword = new TextBox() { Location = new Point(110, 57), Width = 150, PasswordChar = '*', Visible = false };
            this.Controls.Add(txbPassword);

            btnLogin = new Button() { Text = "登入後台", Location = new Point(270, 55), Size = new Size(100, 30), Visible = false };
            btnLogin.Click += new EventHandler(btnLogin_Click);
            this.Controls.Add(btnLogin);

            dataGridView1 = new DataGridView();
            dataGridView1.Location = new Point(20, 110);
            dataGridView1.Size = new Size(740, 420);
            dataGridView1.ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            this.Controls.Add(dataGridView1);
        }

        private void SetupDatabase()
        {
            connstring = new OleDbConnectionStringBuilder();
            connstring.DataSource = dbname;
            connstring.Provider = "Microsoft.Jet.OLEDB.4.0";
        }

        private void LoadEmployeeList()
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
                    cmbEmployee.DataSource = dt;
                    cmbEmployee.DisplayMember = "EmpName";
                    cmbEmployee.ValueMember = "EmpID";
                }
            }
            catch (Exception ex) { MessageBox.Show("載入名單失敗: " + ex.Message); }
        }

        private void rb_CheckedChanged(object sender, EventArgs e)
        {
            if (rbEmployee.Checked)
            {
                lblEmpName.Visible = true; cmbEmployee.Visible = true;
                btnClockIn.Visible = true; btnClockOut.Visible = true;
                lblPwd.Visible = false; txbPassword.Visible = false; btnLogin.Visible = false;
            }
            else
            {
                lblEmpName.Visible = false; cmbEmployee.Visible = false;
                btnClockIn.Visible = false; btnClockOut.Visible = false;
                lblPwd.Visible = true; txbPassword.Visible = true; btnLogin.Visible = true;
            }
        }

        // --- 管理員登入  ---
        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (txbPassword.Text == "1234")
            {
                MessageBox.Show("管理員登入成功！");
                Form2 f2 = new Form2();

                // 當 Form2 關閉時，讓 Form1 重新顯示
                f2.FormClosed += (s, args) => {
                    this.Show();
                    LoadEmployeeList(); // 順便刷新名單
                    ShowLog(); // 順便刷新表格
                    rbEmployee.Checked = true; // 切回員工模式
                    txbPassword.Text = "";
                };

                f2.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("密碼錯誤！");
                txbPassword.Text = "";
            }
        }

        // --- 支援一天多次 ---
        // 輔助類別：用來存最後一筆紀錄的狀態
        struct LastRecordInfo
        {
            public int ID;         // 紀錄編號
            public bool Exists;    // 今天有沒有紀錄
            public bool IsFinished;// 最後一筆是否已下班 (EndTime有值)
        }

        // 取得該員工今天「最後一筆」紀錄的狀態
        private LastRecordInfo GetLastRecordStatus(int empId, string date)
        {
            LastRecordInfo info = new LastRecordInfo { ID = -1, Exists = false, IsFinished = true };

            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
                try
                {
                    conn.Open();
                    // 抓出今天 ID 最大的那一筆 (也就是最新的一筆)
                    string sql = $"SELECT TOP 1 ID, EndTime FROM DailyAttendance WHERE EmpID={empId} AND WorkDate='{date}' ORDER BY ID DESC";
                    OleDbCommand cmd = new OleDbCommand(sql, conn);
                    OleDbDataReader reader = cmd.ExecuteReader();

                    if (reader.Read())
                    {
                        info.Exists = true;
                        info.ID = int.Parse(reader["ID"].ToString());
                        // 如果 EndTime 不是空的，代表已經下班了 (IsFinished = true)
                        // 如果 EndTime 是空的，代表還在工作中 (IsFinished = false)
                        string endTime = reader["EndTime"].ToString();
                        info.IsFinished = !string.IsNullOrEmpty(endTime);
                    }
                }
                catch { }
            }
            return info;
        }

        // --- 上班打卡 ---
        private void btnClockIn_Click(object sender, EventArgs e)
        {
            if (cmbEmployee.SelectedValue == null) { MessageBox.Show("請先選擇員工"); return; }

            int empId = (int)cmbEmployee.SelectedValue;
            string empName = cmbEmployee.Text;
            string today = DateTime.Now.ToString("yyyy/MM/dd");
            string timeNow = DateTime.Now.ToString("HH:mm:ss");

            // 1. 取得狀態
            LastRecordInfo lastInfo = GetLastRecordStatus(empId, today);

            // 2. 判斷邏輯
            // 如果今天沒紀錄 OR 最後一筆已經下班了 -> 可以打上班卡 (新增一筆)
            if (!lastInfo.Exists || lastInfo.IsFinished)
            {
                string sql = $"INSERT INTO DailyAttendance (EmpID, WorkDate, StartTime, Status) VALUES ({empId}, '{today}', '{timeNow}', '工作中')";
                if (TryExecuteSql(sql))
                {
                    MessageBox.Show($"{empName} 上班打卡成功！", "成功");
                    ShowLog();
                }
            }
            else
            {
                // 如果有紀錄且還沒下班
                MessageBox.Show($"您目前還在工作中 (尚未打下班卡)，無法重複打上班卡！");
            }
        }

        // --- 下班打卡 ---
        private void btnClockOut_Click(object sender, EventArgs e)
        {
            if (cmbEmployee.SelectedValue == null) { MessageBox.Show("請先選擇員工"); return; }

            int empId = (int)cmbEmployee.SelectedValue;
            string empName = cmbEmployee.Text;
            string today = DateTime.Now.ToString("yyyy/MM/dd");
            string timeNow = DateTime.Now.ToString("HH:mm:ss");

            // 1. 取得狀態
            LastRecordInfo lastInfo = GetLastRecordStatus(empId, today);

            // 2. 判斷邏輯
            // 如果有紀錄 且 還沒下班 -> 可以打下班卡 (更新那一筆)
            if (lastInfo.Exists && !lastInfo.IsFinished)
            {
                string sql = $"UPDATE DailyAttendance SET EndTime='{timeNow}', Status='已下班' WHERE ID={lastInfo.ID}";
                if (TryExecuteSql(sql))
                {
                    MessageBox.Show($"{empName} 下班打卡成功！", "成功");
                    ShowLog();
                }
            }
            else
            {
                // 如果沒紀錄，或者最後一筆已經下班了
                MessageBox.Show("您目前沒有進行中的上班紀錄，無法打下班卡！\n請先打上班卡。");
            }
        }

        private void ShowLog()
        {
            conn = new OleDbConnection(connStr);
            try
            {
                conn.Open();
                string sql = "SELECT Employee.EmpName AS 姓名, " +
                             "DailyAttendance.WorkDate AS 日期, " +
                             "DailyAttendance.StartTime AS 上班時間, " +
                             "DailyAttendance.EndTime AS 下班時間, " +
                             "DailyAttendance.Status AS 狀態 " +
                             "FROM DailyAttendance, Employee " +
                             "WHERE DailyAttendance.EmpID = Employee.EmpID " +
                             "ORDER BY DailyAttendance.ID DESC";
                OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            catch { }
            finally { conn.Close(); }
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
            catch (Exception ex) { MessageBox.Show("錯誤：" + ex.Message); return false; }
        }
    }
}