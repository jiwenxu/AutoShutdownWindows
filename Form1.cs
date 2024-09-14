using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoShutdown {
    public partial class Form1 : Form {

        private Dictionary<string, List<string>> _holidays = new Dictionary<string, List<string>>();

        public Form1() {
            InitializeComponent();
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e) {
            this.Dispose();
            Application.Exit();
        }

        private void ShowToolStripMenuItem_Click(object sender, EventArgs e) {
            this.Visible = !this.Visible;
            if (this.Visible) {
                this.WindowState = FormWindowState.Normal;
            } else {
                this.WindowState = FormWindowState.Minimized;
            }

        }

        private void Form1_Load(object sender, EventArgs e) {
            this.notifyIcon1.Icon = new Icon(@"window-close.ico");
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.Text = this.Text;
            this.comboBoxHour.SelectedIndex = 23;
            this.comboBoxMinute.SelectedIndex = 0;
            this.textBoxYear.Text = DateTime.Now.Year.ToString();
            this.readHoliday();
            this.selectNextShutdownDay();
            this.monthCalendar1.MinDate = DateTime.Now;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e) {
            e.Cancel = true;
            this.WindowState = FormWindowState.Minimized;
            this.Visible = false;
        }

        private void readHoliday() {
            if (!File.Exists("holiday.json")) {
                return;
            }
            string json = File.ReadAllText("holiday.json");
            this._holidays = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(json);
            this.setBoldedDates();
        }

        private void selectNextShutdownDay() {
            DateTime start = DateTime.Now;
            List<DateTime> boldedDates = new List<DateTime>();
            foreach (string key in this._holidays.Keys) {
                if(int.Parse(key) < start.Year) {
                    continue;
                }

                foreach (string date in _holidays[key]) {
                    DateTime tmp = DateTime.Parse(date);
                    if (tmp < start) {
                        continue;
                    }
                    boldedDates.Add(tmp);
                }
            }
            
            for(int i = 0; i < boldedDates.Count - 1; i++) {
                TimeSpan ts = boldedDates[i + 1] - boldedDates[i];
                if (ts.Days > 1) {
                    this.monthCalendar1.SelectionStart = boldedDates[i];
                    break;
                }
            }
        }

        private void writeHoliday(string json) {
            try {
                StreamWriter sw = new StreamWriter("holiday.json");
                sw.WriteLine(json);
                sw.Close();
            } catch (Exception exception) {
                Console.WriteLine(exception.Message);
            }
        }

        private void setBoldedDates() {
            List<DateTime> boldedDates = new List<DateTime>();
            foreach (string key in this._holidays.Keys) {
                foreach (string date in _holidays[key]) {
                    boldedDates.Add(DateTime.Parse(date));
                }
            }
            this.monthCalendar1.BoldedDates = boldedDates.ToArray();
        }

        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e) {
            if (e.Button == MouseButtons.Right) {
                this.contextMenuStrip1.Show();
            }
        }

        private void button2_Click(object sender, EventArgs e) {
            // 启动文件夹
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            System.Diagnostics.Process.Start(path);
        }

        private void button3_Click(object sender, EventArgs e) {
            // 更新
            string key = this.textBoxYear.Text;
            if (_holidays.ContainsKey(key)) {
                MessageBox.Show("已初始化该年份");
                return;
            }
            DateTime start = DateTime.Parse(key + "-1-1");
            List<string> days = new List<string>();
            while (true) {
                start = start.AddDays(1);
                if (start.DayOfWeek == DayOfWeek.Saturday || start.DayOfWeek == DayOfWeek.Sunday) {
                    days.Add(start.ToString("yyyy-MM-dd"));
                }
                if (start.Year.ToString() != key) {
                    break;
                }
            }
            this._holidays.Add(key, days);
            string json = JsonConvert.SerializeObject(this._holidays);
            writeHoliday(json);
            this.setBoldedDates();
        }

        private void button1_Click(object sender, EventArgs e) {
            //设置关机时间
            DateTime next = DateTime.Parse(this.monthCalendar1.SelectionStart.ToString("yyyy-MM-dd ") + this.comboBoxHour.Text + ":" + this.comboBoxMinute.Text);
            TimeSpan ts = next - DateTime.Now;
            string p = "shutdown /a&shutdown /s /f /t " + (int)ts.TotalSeconds;
            Console.WriteLine(p);
            //this.Exec("shutdown /a");
            this.Exec(p);
            this.textBox1.Text = p;
        }

        private void button4_Click(object sender, EventArgs e) {
            //添加一个休息日
            DateTime date = this.monthCalendar1.SelectionStart;
            string key = date.Year.ToString();
            if (!this._holidays.ContainsKey(key)) {
                MessageBox.Show($"请先始化{key}年");
                return;
            }
            string val = date.ToString("yyyy-MM-dd");
            if (this._holidays[key].Contains(val)) {
                return;
            }
            this._holidays[key].Add(val);
            this._holidays[key].Sort();
            this.writeHoliday(JsonConvert.SerializeObject(this._holidays));
            this.setBoldedDates();
        }

        private void button5_Click(object sender, EventArgs e) {
            // 删除一个休息日
            DateTime date = this.monthCalendar1.SelectionStart;
            string key = date.Year.ToString();
            if (!this._holidays.ContainsKey(key)) {
                MessageBox.Show($"请先始化{key}年");
                return;
            }
            string val = date.ToString("yyyy-MM-dd");
            if (!this._holidays[key].Contains(val)) {
                return;
            }
            this._holidays[key].Remove(val);
            this._holidays[key].Sort();
            this.writeHoliday(JsonConvert.SerializeObject(this._holidays));
            this.setBoldedDates();
        }

        public void Exec(string str) {
            try {
                using (Process process = new Process()) {
                    process.StartInfo.FileName = "cmd.exe";//调用cmd.exe程序
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.RedirectStandardInput = true;//重定向标准输入
                    process.StartInfo.RedirectStandardOutput = true;//重定向标准输出
                    process.StartInfo.RedirectStandardError = true;//重定向标准出错
                    process.StartInfo.CreateNoWindow = true;//不显示黑窗口
                    process.Start();//开始调用执行
                    process.StandardInput.WriteLine(str + "&exit");//标准输入str + "&exit"，相等于在cmd黑窗口输入str + "&exit"
                    process.StandardInput.AutoFlush = true;//刷新缓冲流，执行缓冲区的命令，相当于输入命令之后回车执行
                    process.WaitForExit();//等待退出
                    process.Close();//关闭进程
                    MessageBox.Show("成功");
                }
            } catch {
                MessageBox.Show("失败");
            }
        }

    }
}
