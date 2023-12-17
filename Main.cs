using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using 工作助手.Function;

namespace 格式助手
{
    public partial class Form1 : Form
    {
        DateTime lastCloseButtonClick = DateTime.MinValue;//窗口防误关计时器
        private GlobalKeyboardListener _keyboardListener;//键盘宏初始化
        public Form1()
        {
            InitializeComponent();
            KeyDown += MyForm_KeyDown; // 自定义的KeyDown彩蛋事件处理方法
            FormClosing += new FormClosingEventHandler(Form1_FormClosing);// 自定义的Form1_FormClosing事件处理方法（防误关）
            copyKey.KeyDown += new KeyEventHandler(_copyKey);//键盘宏功能绑定

        }

        private void Form1_Load(object sender, EventArgs e)//开窗时
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);//字符转换界面，默认启动全部功能
            }
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
                toolTip.Show("快速点3下以关闭该应用程序", this, Width / 2 - 50, Height / 2 - 200, 2000);
                e.Cancel = true;
                lastCloseButtonClick = DateTime.Now;
            }
        }

        characterManipulation cm = new characterManipulation();
        string f1; // 将 f1 声明为成员变量
        public void Function1()
        {
            // 第一个功能的实现代码
            f1 = cm.ConvertDatesInText(textBox1.Text);
            textBox2.Text = f1;
        }
        public void Function2()
        {
            // 第二个功能的实现代码
            if (checkedListBox1.GetItemChecked(0))
            {
                // 将分号替换成句号,并清理特殊符号
                string f2 = characterManipulation.RemoveSpecialSymbols(f1, Substitution.Text);
                textBox2.Text = f2;
            }
            else
            {
                textBox2.Text = characterManipulation.RemoveSpecialSymbols(textBox1.Text, Substitution.Text);
            }
        }
        private void Function3()
        {
            // 第三个功能的实现代码
            if (checkedListBox1.GetItemChecked(0) && checkedListBox1.GetItemChecked(1) == false)
            {
                textBox2.Text = characterManipulation.RemoveParentheses(f1);
            }
            else if (checkedListBox1.GetItemChecked(1) && checkedListBox1.GetItemChecked(0) == false)
            {
                textBox2.Text = characterManipulation.RemoveParentheses(characterManipulation.RemoveSpecialSymbols(textBox1.Text, Substitution.Text));
            }
            else if (checkedListBox1.GetItemChecked(0) && checkedListBox1.GetItemChecked(1))
            {
                textBox2.Text = characterManipulation.RemoveParentheses(characterManipulation.RemoveSpecialSymbols(f1, Substitution.Text));
            }
            else
            {
                textBox2.Text = characterManipulation.RemoveParentheses(textBox1.Text);
            }
        }
        //+-书名号和中文引号Start-------------------------
        private void FrenchQuotes_Click(object sender, EventArgs e)
        {
            string inputText = Substitution.Text;
            if (inputText.Contains("《") && inputText.Contains("》"))
            {
                inputText = inputText.Replace("《", "").Replace("》", "");
            }
            else
            {
                inputText = "《》" + inputText;
            }
            Substitution.Text = inputText;
        }

        private void QuotationMark_Click(object sender, EventArgs e)
        {
            string inputText = Substitution.Text;
            if (inputText.Contains("“") && inputText.Contains("”"))
            {
                inputText = inputText.Replace("“", "").Replace("”", "");
            }
            else
            {
                inputText = "“”" + inputText;
            }
            Substitution.Text = inputText;
        }
        //+-书名号和中文引号End-------------------------
        //自动或手动触发文字转换功能Start-------------------------
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            zuoshi(sender, e);
        }
        private void refresh_Click(object sender, EventArgs e)
        {
            zuoshi(sender, e);
        }
        private void zuoshi(object sender, EventArgs e)
        {
            bool anyChecked = false; // 标记是否有多选框被选中
            foreach (object item in checkedListBox1.CheckedItems)
            {
                anyChecked = true; // 至少有一个多选框被选中
                string functionality = item.ToString();
                switch (functionality)
                {
                    case "数字日期转大写":
                        Function1();
                        break;
                    case "特殊符号处理":
                        Function2();
                        break;
                    case "清理括号（包括文本）":
                        Function3();
                        break;
                    default:
                        textBox2.Text = "异常";
                        break;
                }
            }
            // 如果没有任何多选框被选中，执行一个功能
            if (!anyChecked)
            {
                // 执行相应的功能代码
                textBox2.Text = textBox1.Text;
            }
        }
        //自动或手动触发文字转换功能End-------------------------
        //一键写入并获取结果Start-----------------------------
        private void CopyText2_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            if (Clipboard.ContainsText())
            {
                textBox1.Text = Clipboard.GetText();
            }
            else
            {
                textBox1.Text = "剪贴板内没有数据或获取失败";
            }
            Clipboard.SetText(textBox2.Text);
        }
        //一键写入并获取结果End-----------------------------
        //仅复制结果Start-----------------------------
        private void ccopy_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(textBox2.Text);
        }
        //仅复制结果End-----------------------------
        //追加文字Start-----------------------------
        private void zuijia_Click(object sender, EventArgs e)
        {
            textBox1.Text += Clipboard.GetText();
        }
        //追加文字End-----------------------------
        //仅清空按钮Start-----------------------------
        private void set0_Click(object sender, EventArgs e)
        {
            textBox1.Text = null;//仅清空
        }
        //仅清空按钮End-----------------------------
        //压缩包扫描Start-----------------------------
        private void ChooseDir_Click(object sender, EventArgs e)//选目录
        {
            Misc.xuanmulu(dirtext);
        }
        private void run_Click(object sender, EventArgs e)
        {
            string directoryPath = dirtext.Text;

            ZipScanner zipScanner = new ZipScanner();
            zipScanner.ScanDirectory(directoryPath, CheckTxtDataBox.Checked);

            jieguo.Text = zipScanner.GetOutput();
        }
        private void Del_Click(object sender, EventArgs e)
        {
            jieguo.Text = "检测结果将会出现在这里。当前状态：空闲";
        }
        //压缩包扫描End-----------------------------

        //表格扫描Start-----------------------------
        private async void RunXls_Click(object sender, EventArgs e)//扫表格
        {
            btn_stopScanXls.Enabled = true;
            runXls.Enabled = false;
            string extension = Path.GetExtension(xlsDir.Text);

            if (extension == ".xls" || extension == ".xlsx" || extension == ".csv")
            {

                textBox5.Text = "扫描中，稍等。。。\r\n";
                ScanXls scanXls = new ScanXls();
                progressBar1.Maximum = (int)numericUpDown2.Value;

                await scanXls.CheckExcelLinks(xlsDir.Text,
                    (int)numericUpDown1.Value,
                    (int)numericUpDown2.Value,
                    (int)numericUpDown3.Value,
                    yellow.Checked, progressBar1);

                textBox5.Text += scanXls.GetOutput();

            }
            else
            {
                textBox5.Text = "请选择一个表格再开始。（路径或文件异常）";
            }
            btn_stopScanXls.Enabled = false;
            runXls.Enabled = true;
        }

        private void xuan_Click(object sender, EventArgs e)//选表格
        {
            xlsDir.Text= Misc.xuanbiaoge();
        }
        private void clear_Xls_Click(object sender, EventArgs e)//清除按钮
        {
            textBox5.Text = "检测结果将会出现在这里。当前状态：空闲\r\n目前可以检测链接是否能正常打开";
        }
        private void btn_stopScanXls_Click(object sender, EventArgs e)//停止扫描
        {
            ScanXls.cancellationTokenSource.Cancel();
        }
        //表格扫描End-----------------------------


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
            if (comboBox1.SelectedIndex == 3) 
            {
                string biaoge = Misc.xuanbiaoge();
                List<string> dir= new List<string>();
                dir.Add(dir1.Text);
                dir.Add(dir2.Text);
                scanjieguo2.Text = "选择的表格是："+ biaoge;
                if(biaoge != "")
                scanZip2.CompareDataWithZipFile(biaoge, "A", dir);
                scanjieguo2.Text = scanZip2.GetOutput();
            }
            else if (comboBox1.SelectedIndex != -1)
            {
                
                scanZip2.CheckZip(dir1.Text, dir2.Text, comboBox1.SelectedIndex);
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

        //Win7文本框全选兼容性代码Start-----------------------
        private void textBox2_KeyDown(object sender, KeyEventArgs e)//win7兼容性，ctrl+a全选结果
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                textBox2.SelectAll();
                e.SuppressKeyPress = true; // 防止同时触发默认的全选操作
            }
        }
        private void textBox1_KeyDown(object sender, KeyEventArgs e)//win7兼容性，ctrl+a全选结果
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                textBox1.SelectAll();
                e.SuppressKeyPress = true; // 防止同时触发默认的全选操作
            }
        }
        //Win7文本框全选兼容性代码End-------------------------------
        

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
        //拖动放置文件目录End------------------------------

    }
}
