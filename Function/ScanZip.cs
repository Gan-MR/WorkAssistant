using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace 工作助手.Function
{
    public class ZipScanner
    {

        private StringBuilder outputBuilder;
        int bugjisu = 0;
        int zipjisu = 0;
        public ZipScanner()
        {
            outputBuilder = new StringBuilder();
        }
        private void AppendOutput(string text)
        {
            outputBuilder.AppendLine(text);
            bugjisu++;
        }
        public string GetOutput()
        {
            return outputBuilder.ToString();
        }

        public void ScanDirectory(string directoryPath)
        {
            try
            {
                string[] subDirectories = Directory.GetDirectories(directoryPath);

                if (subDirectories.Length == 0)
                {
                    AppendOutput("该目录下没有子文件夹");
                    return;
                }

                foreach (string subDirectory in subDirectories)
                {
                    // 执行扫描操作
                    CheckImageCountInDirectory(subDirectory); // 检查文件夹中图片数量是否符合规格
                    zipjisu++;
                }

                if (GetOutput().Length == 0)
                {
                    AppendOutput("扫描已完成，扫描了 " + zipjisu + " 个子文件夹，没有发现已知问题");
                }
                else
                {
                    AppendOutput("本次扫描了 " + zipjisu + " 个子文件夹，发现了 " + bugjisu + " 个问题");
                }
            }
            catch (Exception)
            {
                AppendOutput("扫描启动失败，未设定目录或选择所目录异常");
            }
        }

        public void CheckImageCountInDirectory(string directoryPath)
        {
            try
            {
                string[] imageFiles = Directory.GetFiles(directoryPath, "*.jpg");

                int imageCount = imageFiles.Length;

                if (imageCount < 3)
                {
                    AppendOutput(Path.GetFileName(directoryPath) + " - 文件夹中包含的.jpg图片少于3个。");
                }
                else if (imageCount > 5)
                {
                    AppendOutput(Path.GetFileName(directoryPath) + " - 文件夹中包含的.jpg图片多于5个。");
                }

            }
            catch (Exception)
            {
                AppendOutput("扫描文件夹 " + Path.GetFileName(directoryPath) + " 时出现错误：可能该文件夹没有符合规格的图片文件或者没有访问权限");
            }
        }








    }
}