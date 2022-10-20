using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;


using System.Runtime.InteropServices;
using System.Drawing.Text;
using System.IO;

using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;


using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using Keys = System.Windows.Forms.Keys;

namespace E_COUNT_MSGR2

{


    public partial class Form1 : Form
    {
        Thread thread = null;
        Point fPt;
        bool isMove;


        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateRoundRectRgn(int x1, int y1, int x2, int y2,
                                                int cx, int cy);
        [DllImport("user32.dll")]
        private static extern int SetWindowRgn(IntPtr hWnd, IntPtr hRgn, bool bRedraw);

        [DllImport("user32.dll")]
        private static extern int ShowWindow(IntPtr windowHandle, int command);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);


        int quit_code = 0;
        int open_msgr = 0;

        public Form1()
        {
            InitializeComponent();
            msgr.ContextMenuStrip = contextMenuStrip1;

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            IntPtr ip = CreateRoundRectRgn(0, 0, button3.Width, button3.Height, 15, 15);
            int i = SetWindowRgn(button3.Handle, ip, true);

            StreamReader sr;
            sr = new StreamReader(System.Environment.CurrentDirectory + @"\set.dat", Encoding.UTF8);

            string line;
            string set_data = "";

            while ((line = sr.ReadLine()) != null)
            {
                set_data += line + Environment.NewLine;
            }
            sr.Close();

            string[] setting = set_data.Split('\n');

            Console.WriteLine(setting[0]);

            if (setting[0].Trim() == "ok")
            {
                checkBox1.Checked = true;
                textBox1.Text = setting[1].Trim();
                textBox2.Text = setting[2].Trim();
                textBox3.PasswordChar = '*';
                textBox3.Text = setting[3].Trim();
            }
            else
            {
                checkBox1.Checked = false;
            }

            
            new DriverManager().SetUpDriver(new ChromeConfig());

          

            
        }

     


        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            isMove = true;
            fPt = new Point(e.X, e.Y);
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (isMove && (e.Button & MouseButtons.Left) == MouseButtons.Left)
                Location = new Point(this.Left - (fPt.X - e.X), this.Top - (fPt.Y - e.Y));
        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            isMove = false;
        }

        private void save_set()
        {
            StreamWriter sw;
            sw = new StreamWriter(System.Environment.CurrentDirectory + @"\set.dat");

            if (checkBox1.Checked)
            {
                string remember_stat = "ok";
                string cp_name = textBox1.Text.ToString();
                string usr_id = textBox2.Text.ToString();
                string pwd = textBox3.Text.ToString();

                sw.WriteLine(remember_stat);
                sw.WriteLine(cp_name);
                sw.WriteLine(usr_id);
                sw.WriteLine(pwd);
            }

            else
            {
                string remember_stat = "no";
                sw.WriteLine(remember_stat);
            }
            sw.Close();
        }



        private void hide_window()

        {
            int SW_HIDE = 0;
            IntPtr hWnd = FindWindow(null, "이카운트 로그인 | ECOUNT ERP - Chrome");
            ShowWindow(hWnd, SW_HIDE);
        }

        
        public void get_()
        {
            int login_check = 0;
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;

            var options = new ChromeOptions();
            options.EnableMobileEmulation("Galaxy S5");
            options.AddExcludedArgument("enable-automation");
            options.AddAdditionalCapability("useAutomationExtension", false);
            
            var driver = new ChromeDriver(driverService, options);
            

            

            driver.Navigate().GoToUrl("https://login.ecount.com/Login/");
            hide_window();

            while (true)
            {
                try
                {
                    driver.FindElementByXPath("//*[@id='com_code']").Click();
                    driver.FindElementByXPath("//*[@id='com_code']").SendKeys(this.textBox1.Text.ToString());
                    break;
                }

                catch (Exception ex)
                {
                    continue;
                }
            }

            driver.FindElementByXPath("//*[@id='id']").Click();
            driver.FindElementByXPath("//*[@id='id']").SendKeys(this.textBox2.Text.ToString());

            driver.FindElementByXPath("//*[@id='passwd']").Click();
            driver.FindElementByXPath("//*[@id='passwd']").SendKeys(this.textBox3.Text.ToString());

            driver.FindElementByXPath("//*[@id='save']").Click();

            while (true)
            {
                try
                {
                    var msgr = driver.FindElementByXPath("//*[@id='ecMessenger']");
                    driver.ExecuteScript("arguments[0].click()", msgr);
                    break;
                }

                catch (Exception ex)
                {
                    Console.WriteLine("오류:" + ex.Message);
                    if (ex.Message.Contains("no such element"))
                    {
                        Console.WriteLine("찾을 수 없음");
                        MessageBox.Show("ID, PW 오류", "로그인 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        login_check = 1;
                        break;
                    }

                    continue;
                }
            }

            if (login_check == 0)
            {
                while (true)
                {
                    if (quit_code == 1)
                    {
                        Console.WriteLine("cliked quit");
                        driver.Quit();
                        Application.Exit();
                        break;
                    }
                    Thread.Sleep(1000);

                    //IntPtr hWnd = FindWindow(null, "ECMessenger - Chrome");
                    //Console.WriteLine("ecount: " + hWnd);

                    //if (hWnd == IntPtr.Zero)
                    //{
                    //    driver.Quit();
                    //    Application.Exit();
                    //    break;
                    //}

                    if (open_msgr == 1)
                    {
                        Console.WriteLine("cliked open");
                        var msgr = driver.FindElementByXPath("//*[@id='ecMessenger']");
                        driver.ExecuteScript("arguments[0].click()", msgr);
                        open_msgr = 0;
                    }
                }
            }

            if (login_check == 1)
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new MethodInvoker(delegate { this.Visible = true; }));
                }
                thread.Abort();
                thread.Join();
            }

            //driver.ExecuteScript("document.title = 'msgr - chrome'");

        }


        private void ProcessKill(string imgName)
        {
            // 프로세스 종료
            Process[] processes = Process.GetProcessesByName(imgName);
            foreach (var process in processes)
            {
                try
                {
                    process.Kill();
                }
                catch (Exception e) { }
            }
        }




        private void button3_Click(object sender, EventArgs e)
        {
            save_set();

            thread = new Thread(new ThreadStart(get_));
            thread.IsBackground = true;
            thread.Start();

            this.Visible = false;

           
        }

        
       public void isChecked(bool Checked)
        {
            if (Checked)
            {
                checkBox1.Image = Properties.Resources.ON_2;
            }

            else
            {
                checkBox1.Image = Properties.Resources.OFF_2;
            }
        }


        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            isChecked(checkBox1.Checked);
        }

        private void textBox1_Click(object sender, EventArgs e)
        {

            if (textBox1.Text == "COMPANY ID")
            {
                textBox1.Text = String.Empty;
            }
            
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == "USER ID")
            {
                textBox2.Text = String.Empty;
            }
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "PASSWORD")
            {
                textBox3.Text = String.Empty;
                textBox3.PasswordChar = '*';
            }
            
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox3.PasswordChar = '*';
        }


        private void 종료ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            quit_code = 1;
            //ProcessKill("chromedriver.exe");
            //ProcessKill("E-COUNT_MSGR2.exe");
            //Application.Exit();
        }

        private void 열기ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Console.WriteLine("open");
            open_msgr = 1;
        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                button3_Click(sender, e);
            }
        }
    }
}
