using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using 工作助手.Function;

namespace 格式助手
{
    public partial class Form1 : Form
    {
        DateTime lastCloseButtonClick = DateTime.MinValue;//窗口防误关计时器
        private GlobalKeyboardListener _keyboardListener;//键盘宏初始化
        string Scanzip0String;//用来保存结果输出框中的说明文字。
        public Form1()
        {
            InitializeComponent();
            KeyDown += MyForm_KeyDown; // 自定义的KeyDown彩蛋事件处理方法
            FormClosing += new FormClosingEventHandler(Form1_FormClosing);// 自定义的Form1_FormClosing事件处理方法（防误关）
            copyKey.KeyDown += new KeyEventHandler(_copyKey);//键盘宏功能绑定
            
            Scanzip0String = jieguo.Text;//启动时保存一次
        }

        private void Form1_Load(object sender, EventArgs e)//开窗时
        {
            
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)//关窗（防误关）
        {
            if ((DateTime.Now - lastCloseButtonClick).TotalMilliseconds < SystemInformation.DoubleClickTime)
            {
                // 用户在短时间内两次点击了关闭按钮，执行关闭操作
                e.Cancel = false;
            }
            else
            {
                // 用户第一次点击关闭按钮，弹出悬浮提示
                ToolTip toolTip = new ToolTip();
                toolTip.Show("快速点3下以关闭该应用程序", this, Width / 2 + 100, Height / 2-323 , 2000);
                e.Cancel = true;
                lastCloseButtonClick = DateTime.Now;
            }
        }

        
        //扫描Start-----------------------------
        private void ChooseDir_Click(object sender, EventArgs e)//选目录
        {
            Misc.xuanmulu(dirtext);
        }
        private void run_Click(object sender, EventArgs e)
        {
            string directoryPath = dirtext.Text;

            ZipScanner zipScanner = new ZipScanner();
            zipScanner.ScanDirectory(directoryPath);

            jieguo.Text = zipScanner.GetOutput();
        }

        
        private void Del_Click(object sender, EventArgs e)
        {
            jieguo.Text = Scanzip0String;
            
        }
        //扫描End-----------------------------

        



        //批量压缩包校验Start-----------------------------
        private void xuan1_Click(object sender, EventArgs e)
        {
            Misc.xuanmulu(dir1);
        }

        private void xuan2_Click(object sender, EventArgs e)
        {
            Misc.xuanmulu(dir2);
        }

        private void runScan_Click(object sender, EventArgs e)
        {
            ScanZip2 scanZip2 = new ScanZip2();
            if (comboBox1.SelectedIndex == 1)
            {
                string biaoge = Misc.xuanbiaoge();
                List<string> dir = new List<string>();
                dir.Add(dir1.Text);
                dir.Add(dir2.Text);
                scanjieguo2.Text = "选择的表格是：" + biaoge + "\r\n";
                if (biaoge != "")
                    scanZip2.CompareDataWithFolders(biaoge, "A", dir);
                scanjieguo2.Text += scanZip2.GetOutput();
            }
            else if (comboBox1.SelectedIndex == 0)
            {

                scanZip2.CheckZip(dir1.Text, dir2.Text);
                scanjieguo2.Text = scanZip2.GetOutput();
            }
            else
            {
                scanjieguo2.Text = "没有选择功能";
            }

        }

        private void qingli_Click(object sender, EventArgs e)//清除结果
        {
            scanjieguo2.Text = "检测结果将会出现在这里。当前状态：空闲";
        }
        //批量压缩包校验End-----------------------------
        //键盘宏Start----------------------
        bool MagicT = true;

        private void Magic_Click(object sender, EventArgs e)
        {
            if (MagicT)
            {
                MagicT = false;
                _keyboardListener = new GlobalKeyboardListener();
                _keyboardListener.KeyDown += Misc.KeyboardListener_KeyDown;
                _keyboardListener.Start();
                MagicZ.Text = "运行中";
                tabPage6.Text = "键盘宏（运行中）";
                MagicZ.ForeColor = Color.Green;
            }
            else
            {
                MagicT = true;
                _keyboardListener.Stop();
                MagicZ.Text = "未运行";
                tabPage6.Text = "键盘宏";
                MagicZ.ForeColor = Color.Red;
            }

        }
        private void _copyKey(object sender, KeyEventArgs e)
        {
            Keys pressedKey = e.KeyCode;
            KeysConverter converter = new KeysConverter();
            copyKey.Text = "";
            copyKey.Text = converter.ConvertToString(pressedKey);
        }

        //键盘宏End-----------------------

        


        //彩蛋
        private void MyForm_KeyDown(object sender, KeyEventArgs e)//彩蛋
        {
            if (e.KeyCode == Keys.Escape) // 如果按下的是Esc键
            {
                Text = Text.Replace("工作助手", "下班加速器");
            }
        }

        //拖动放置文件目录Start------------------------------
        private void dirtext_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && ((e.AllowedEffect & DragDropEffects.Copy) == DragDropEffects.Copy))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void dirtext_DragDrop(object sender, DragEventArgs e)
        {
            // 获取拖放的文件夹路径
            string[] folders = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (folders.Length > 0 && Directory.Exists(folders[0]))
            {
                dirtext.Text = folders[0];
            }
        }

        private void dir1_DragDrop(object sender, DragEventArgs e)
        {
            // 获取拖放的文件夹路径
            string[] folders = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (folders.Length > 0 && Directory.Exists(folders[0]))
            {
                dir1.Text = folders[0];
            }
        }

        private void dir1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && ((e.AllowedEffect & DragDropEffects.Copy) == DragDropEffects.Copy))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void dir2_DragDrop(object sender, DragEventArgs e)
        {
            // 获取拖放的文件夹路径
            string[] folders = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (folders.Length > 0 && Directory.Exists(folders[0]))
            {
                dir2.Text = folders[0];
            }
        }

        private void dir2_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && ((e.AllowedEffect & DragDropEffects.Copy) == DragDropEffects.Copy))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        private void FileDir_DragDrop(object sender, DragEventArgs e)
        {
            // 获取拖放的文件夹路径
            string[] folders = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (folders.Length > 0 && Directory.Exists(folders[0]))
            {
                FileDir.Text = folders[0];
            }
        }
        private void FileDir_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && ((e.AllowedEffect & DragDropEffects.Copy) == DragDropEffects.Copy))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        //拖动放置文件目录End------------------------------



        private void FileXuan_Click(object sender, EventArgs e)
        {
            Misc.xuanmulu(FileDir);
        }

        private void oneDesktop_Click(object sender, EventArgs e)
        {
            FileDir.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }
        private void nulldir1_Click(object sender, EventArgs e)
        {
            FileStar.Text = "";
        }
        
        private void generate_Click(object sender, EventArgs e)
        {
            if (Regex.IsMatch(FileStar.Text, @"\d")) {
                if (Directory.Exists(FileDir.Text)&& FileDir.Text!="") { 
                    Misc.GenerateFiles(FileStar.Text, (int)numFile.Value, FileDir.Text,wenjianjia.Checked);
                    MessageBox.Show("生成完毕", "程序没有爆", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("无效的放置目录", "程序爆了", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
            } else {
                MessageBox.Show("初始文件名中至少包含一个数字", "程序爆了", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
                
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            FileDir.Text = "";
        }


        private void update_Click(object sender, EventArgs e)
        {
            UpDate.InstallUpdateSyncWithInfo();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
