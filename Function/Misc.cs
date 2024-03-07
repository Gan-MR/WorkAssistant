using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace 工作助手.Function
{
    class Misc
    {
        public static void KeyboardListener_KeyDown(object sender, KeyEventArgs e)//键盘宏实现
        {
            if (e.KeyCode == Keys.F1)
            {
                // 执行复制操作
                SendKeys.Send("^c");
                e.Handled = true; // 取消键的默认行为
            }
            else if (e.KeyCode == Keys.F2) // win键
            {
                SendKeys.Send("^v");
                // 执行粘贴操作
                e.Handled = true; // 取消键的默认行为
            }
            else if (e.KeyCode == Keys.F3)
            {
                SendKeys.Send("^a");
                // 执行粘贴操作
                e.Handled = true; // 取消键的默认行为
            }
            else if (e.KeyCode == Keys.F4)
            {
                SendKeys.Send("^x");
                // 执行粘贴操作
                e.Handled = true; // 取消键的默认行为
            }




        }


        public static void xuanmulu(TextBox textBox)//选择文件目录
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "选择压缩包所在的目录";
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                string path = fbd.SelectedPath;
                textBox.Text = path;
            }
        }

        public static string xuanbiaoge()//选择文件目录
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
                    return selectedFilePath;
                }
            }
            return "";
        }


        public static void GenerateFiles(string text, int number, string filePath,bool dirOrTxt)
        {
            string prefix = new string(text.TakeWhile(c => !char.IsDigit(c)).ToArray());
            string suffix = new string(text.Reverse().TakeWhile(c => !char.IsDigit(c)).Reverse().ToArray());
            string digits = new string(text.Where(char.IsDigit).ToArray());
            int originalNumber = int.Parse(digits);
            if (dirOrTxt)
            {
                for (int i = 0; i < number; i++)
                {
                    string newNumber = (originalNumber + i).ToString();
                    string folderName = prefix + newNumber + suffix;
                    Directory.CreateDirectory(Path.Combine(filePath, folderName));
                }
            }
            else
            {
                for (int i = 0; i < number; i++)
                {
                    string newNumber = (originalNumber + i).ToString();
                    string fileName = prefix + newNumber + suffix + ".txt";
                    File.Create(Path.Combine(filePath, fileName)).Dispose();

                }
            }
            
        }



    }
}
