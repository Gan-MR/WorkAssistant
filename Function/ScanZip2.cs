using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace 工作助手.Function
{
    public class ScanZip2
    {
        private StringBuilder outputBuilder;
        public ScanZip2()
        {
            outputBuilder = new StringBuilder();
        }
        private void AppendOutput(string text)
        {
            outputBuilder.AppendLine(text);

        }
        public string GetOutput()
        {
            return outputBuilder.ToString();
        }
        public void CheckZip(string dir1, string dir2, int functionIndex)
        {
            if (functionIndex == 0)
            {
                if (Directory.Exists(dir1))
                {
                    CheckZipFiles(dir1);
                }
                else
                {
                    AppendOutput("目录1：找不到对象");
                }

            }
            else if (functionIndex == 1)
            {
                if (Directory.Exists(dir1) && Directory.Exists(dir2))
                { CompareZipFiles(dir1, dir2); }
                else
                {
                    AppendOutput("目录1或目录2：找不到这个位置");
                }

            }
            else if (functionIndex == 2)
            {
                if (Directory.Exists(dir1) && Directory.Exists(dir2))
                { CompareZipFilesInDirectories(dir1, dir2); }
                else
                {
                    AppendOutput("目录1或目录2：找不到这个位置");
                }
            }

        }


        public void CheckZipFiles(string directoryPath)
        {

            // 获取目录下所有压缩包文件
            string[] zipFiles = Directory.GetFiles(directoryPath, "*.zip");

            // 提取所有压缩包文件名中的数字部分
            List<int> numbers = new List<int>();
            foreach (string zipFile in zipFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(zipFile);
                Match match = Regex.Match(fileName, @"\d+");
                if (match.Success)
                {
                    int number = int.Parse(match.Value);
                    numbers.Add(number);
                }
            }

            // 校验数字是否连续
            numbers.Sort();//排序
            int missingCount = 0;
            for (int i = 1; i < numbers.Count; i++)
            {
                if (numbers[i] != numbers[i - 1] + 1)
                {
                    // 输出附近的1个压缩包文件名
                    int prevNumber = numbers[i - 1];
                    int nextNumber = numbers[i];
                    string prevZipFile = zipFiles.First(f => f.Contains(prevNumber.ToString()));
                    string nextZipFile = zipFiles.First(f => f.Contains(nextNumber.ToString()));

                    int count = numbers[i] - numbers[i - 1] - 1;
                    missingCount += count;
                    if (count > 70) { AppendOutput("区间异常"); }
                    else
                    {
                        AppendOutput("在 " + Path.GetFileName(prevZipFile) + " 和 " + Path.GetFileName(nextZipFile) + " 之间，缺少 " + count + " 个压缩包");
                    }
                }
            }
            if (missingCount != 0)
            {
                AppendOutput("总共缺少 " + missingCount + " 个压缩包。");
            }
            else
            {
                AppendOutput("没有找到不连续的压缩包。");
            }

        }


        public void CompareZipFiles(string directoryPath1, string directoryPath2)
        {
            // 获取两个目录下的压缩包文件
            string[] zipFiles1 = Directory.GetFiles(directoryPath1, "*.zip");
            string[] zipFiles2 = Directory.GetFiles(directoryPath2, "*.zip");

            // 比对两个目录下的压缩包数量
            if (zipFiles1.Length != zipFiles2.Length)
            {
                if (zipFiles1.Length > zipFiles2.Length)
                {
                    AppendOutput("目录1的文件比目录2多");
                }
                else if (zipFiles1.Length < zipFiles2.Length)
                {
                    AppendOutput("目录2的文件比目录1多");
                }
                else
                {
                    AppendOutput("目录1的文件数量与目录2一致");
                }
            }

            // 检测两个目录下的所有压缩包的文件名是否一致
            List<string> fileNames1 = zipFiles1.Select(Path.GetFileName).ToList();
            List<string> fileNames2 = zipFiles2.Select(Path.GetFileName).ToList();

            foreach (string fileName in fileNames1)
            {
                if (!fileNames2.Contains(fileName))
                {
                    AppendOutput("压缩包 " + fileName + " 存在目录1中，目录2没有");
                }
            }

            foreach (string fileName in fileNames2)
            {
                if (!fileNames1.Contains(fileName))
                {
                    AppendOutput("压缩包 " + fileName + " 存在目录2中，目录1没有");
                }
            }
        }



        public void CompareZipFilesInDirectories(string directory1, string directory2)
        {
            // 获取目录1下缺少的压缩包

            List<int> missingNumbers = new List<int>();
            string[] zipFilesInDirectory1 = Directory.GetFiles(directory1, "*.zip");
            string firstLetter = "";
            string name = "";
            List<string> missingZipFileNames = new List<string>();
            foreach (string zipFile in zipFilesInDirectory1)
            {
                string fileName = Path.GetFileNameWithoutExtension(zipFile);
                Match match = Regex.Match(fileName, @"\d+");
                if (match.Success)
                {
                    int number = int.Parse(match.Value);
                    missingNumbers.Add(number);
                }
            }

            Match match1 = Regex.Match(Path.GetFileNameWithoutExtension(zipFilesInDirectory1[0]), @"[a-zA-Z]");
            if (match1.Success) { firstLetter = match1.Value[0].ToString(); }
            MatchCollection matches = Regex.Matches(Path.GetFileNameWithoutExtension(zipFilesInDirectory1[0]), @"[\u4e00-\u9fff]"); // 匹配中文字符的 Unicode 范围
            foreach (Match match in matches) { name += match.Value; }

            missingNumbers.Sort();//排序
            for (int i = 1; i < missingNumbers.Count; i++)
            {
                for (int j = missingNumbers[i - 1] + 1; j < missingNumbers[i]; j++)//提取断开的数字
                {
                    string zipFileName = firstLetter + j + name + ".zip";
                    missingZipFileNames.Add(zipFileName);
                }

            }

            // 在目录2下比对是否存在这些压缩包
            string[] zipFilesInDirectory2 = Directory.GetFiles(directory2, "*.zip");
            List<string> existingZipFileNames = new List<string>(zipFilesInDirectory2.Select(Path.GetFileName));

            foreach (string missingZipFileName in missingZipFileNames)
            {
                if (existingZipFileNames.Contains(missingZipFileName))
                {
                    AppendOutput("目录1缺失的" + missingZipFileName + "存在于目录2");
                }
                else
                {
                    AppendOutput("目录2下不存在压缩包：" + missingZipFileName);
                }
            }
        }









    }
}

