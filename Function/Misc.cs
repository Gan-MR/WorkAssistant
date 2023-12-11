using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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









    }
}
