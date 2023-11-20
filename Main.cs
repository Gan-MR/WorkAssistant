using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using 工作助手.Function;

namespace 格式助手
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            KeyPreview = true; // 允许窗体接收键盘事件
            KeyDown += MyForm_KeyDown; // 注册KeyDown事件处理方法
        }
        


        private void Form1_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);//字符转换界面，默认启动全部功能
            }

            jieguo.Font = new Font("Microsoft YaHei", 12, FontStyle.Regular, GraphicsUnit.Point, ((byte)(134)));
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
                string f2 = characterManipulation.RemoveSpecialSymbols(f1,Substitution.Text);
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
        private void ChooseDir_Click(object sender, EventArgs e)
        {
            string path = string.Empty;
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "选择压缩包所在的目录";
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                path = fbd.SelectedPath;
                dirtext.Text = path;
            }
        }
        private void run_Click(object sender, EventArgs e)
        {
            string directoryPath = dirtext.Text;

            ZipScanner zipScanner = new ZipScanner();
            zipScanner.ScanDirectory(directoryPath);

            jieguo.Text = zipScanner.GetOutput();
        }
        private void del_Click(object sender, EventArgs e)
        {
            jieguo.Text = "检测结果将会出现在这里。当前状态：空闲";
        }
        //压缩包扫描End-----------------------------



        //表格扫描Start-----------------------------
        private void runXls_Click(object sender, EventArgs e)//扫表格
        {
            string extension = System.IO.Path.GetExtension(xlsDir.Text);

            if (extension == ".xls" || extension == ".xlsx" || extension == ".csv")
            {

                textBox5.Text = "扫描中，稍等。。。";
                ScanXls scanXls = new ScanXls();
                scanXls.CheckExcelLinks(xlsDir.Text, 
                    (int)numericUpDown1.Value,
                    (int)numericUpDown2.Value, 
                    (int)numericUpDown3.Value,
                    yellow.Checked);
                textBox5.Text +=scanXls.GetOutput();
            }
            else
            {
                textBox5.Text = "请选择一个表格再开始。";
            }
        }
        
        private void xuan_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                openFileDialog.Filter = "Excel文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFilePath = openFileDialog.FileName;
                    // 在这里处理选择的文件路径
                    xlsDir.Text = selectedFilePath;
                }
            }
        }
        private void qing_Click(object sender, EventArgs e)
        {
            textBox5.Text = "检测结果将会出现在这里。当前状态：空闲";
        }
        //表格扫描End-----------------------------
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



        //更多功能。。。。。。。。。

        //彩蛋
        private void MyForm_KeyDown(object sender, KeyEventArgs e)//彩蛋
        {
            if (e.KeyCode == Keys.Escape) // 如果按下的是Esc键
            {
                Text = Text.Replace("工作助手", "下班加速器");
            }
        }

        
    }
}
