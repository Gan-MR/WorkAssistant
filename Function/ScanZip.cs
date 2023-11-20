using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows.Forms;




namespace 工作助手.Function
{
    public class ZipScanner
    {

        private StringBuilder outputBuilder;
        int jisu = 0;
        int zipjisu = 0;
        public ZipScanner()
        {
            outputBuilder = new StringBuilder();
        }

        public void ScanDirectory(string directoryPath)
        {
            
            string[] zipFiles = Directory.GetFiles(directoryPath, "*.zip");

            if (zipFiles.Length == 0)
            {
                AppendOutput("该目录下没有找到.zip压缩包文件");
                return;
            }

            foreach (string zipFile in zipFiles)
            {
                // 执行扫描操作
                
                CheckTextFileCountAndImageCount(zipFile);//文本=1，且文本行数=图片张数，文本第一行少于四个字
                CheckTextFileContent(zipFile);//压缩包内文本超过17个连续字符相同字符，文本是否为空，是否包含空格，笔者，小编等关键字
                CheckTextFileNumber(zipFile);//文本的数字与压缩包的数字
                CheckNestedZipFiles(zipFile);//嵌套压缩包
                zipjisu++;
            }
            if (GetOutput().Length == 0)
            {

                AppendOutput("扫描已完成，扫描了"+ zipjisu +"个压缩包，没有发现已知问题");
                
            }
            else
            {
                AppendOutput("本次扫描了 "+zipjisu+" 个压缩包，发现了 "+ jisu + " 个问题");
            }

        }

        

        private void CheckTextFileCountAndImageCount(string zipFile)//检测文本=1，且文本行数=图片张数，文本第一行不少于四个字
        {
            int textFileCount = 0;
            int imageCount = 0;
            int textLineCount = 1;
            bool textFileFound = false;
            bool textLineAndImageCountMismatch = false;

            using (ZipArchive archive = ZipFile.OpenRead(zipFile))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                    {
                        textFileCount++;
                        textFileFound = true;
                        using (StreamReader reader = new StreamReader(entry.Open()))
                        {
                            string firstLine = reader.ReadLine();
                            if (firstLine != null && firstLine.Length < 4)
                            {
                                AppendOutput($"{Path.GetFileName(zipFile)} - .txt文本的第一行小于4个字");
                            }

                            while (reader.ReadLine() != null)
                            {
                                textLineCount++;
                            }
                        }
                    }
                    else if (entry.FullName.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase))
                    {
                        imageCount++;
                    }
                }
            }

            if (!textFileFound)
            {
                AppendOutput($"{Path.GetFileName(zipFile)} - 未找到文本文件");
            }
            else if (textFileCount == 1 && textLineCount == imageCount)
            {
                // Do something if the conditions are met
            }
            else
            {
                textLineAndImageCountMismatch = true;
            }

            if (textLineAndImageCountMismatch)
            {
                AppendOutput($"{Path.GetFileName(zipFile)} - 文本行数和图片数量不一致，文本行数: {textLineCount} ，jpg图片数量: {imageCount}");
            }
        }



        private void CheckTextFileContent(string zipFile)//压缩包内文本超过17个连续字符相同字符，文本是否为空，是否包含空格
        {
            List<string> repeatedLines = new List<string>(); // 用于存储重复的文本行
            Dictionary<string, List<int>> lineIndexMap = new Dictionary<string, List<int>>(); // 用于存储每个文本行的出现位置

            using (ZipArchive archive = ZipFile.OpenRead(zipFile))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                    {
                        using (StreamReader reader = new StreamReader(entry.Open(), Encoding.Default))
                        {
                            int lineIndex = 0;
                            string line;

                            bool isEmpty = true; // 用于标记文本文件是否为空

                            while ((line = reader.ReadLine()) != null)
                            {
                                lineIndex++;
                                isEmpty = false; // 当读取到文本行时，将 isEmpty 标记为 false

                                if (line.Length > 17)
                                {
                                    if (lineIndexMap.ContainsKey(line))
                                    {
                                        lineIndexMap[line].Add(lineIndex);
                                    }
                                    else
                                    {
                                        lineIndexMap[line] = new List<int> { lineIndex };
                                    }
                                }

                                // 检测文本中是否包含空格
                                if (line.Contains(" "))
                                {
                                    int startIndex = Math.Max(0, line.IndexOf(" ") - 8);
                                    int endIndex = Math.Min(line.Length, line.IndexOf(" ") + 8 + 1);
                                    string context = line.Substring(startIndex, endIndex - startIndex);
                                    AppendOutput($"{Path.GetFileName(zipFile)} - 该压缩包文本的第 { lineIndex } 行包含空格: {context}");
                                }
                                if (line.Contains("小编"))// 检测文本中是否包含“小编”
                                {
                                    int startIndex = Math.Max(0, line.IndexOf("小编") - 8);
                                    int endIndex = Math.Min(line.Length, line.IndexOf("小编") + 8 + 1);
                                    string context = line.Substring(startIndex, endIndex - startIndex);
                                    AppendOutput($"{Path.GetFileName(zipFile)} - 该压缩包文本的第 { lineIndex } 行包含“小编”: {context}");
                                }
                                if (line.Contains("笔者"))// 检测文本中是否包含“笔者”
                                {
                                    int startIndex = Math.Max(0, line.IndexOf("笔者") - 8);
                                    int endIndex = Math.Min(line.Length, line.IndexOf("笔者") + 8 + 1);
                                    string context = line.Substring(startIndex, endIndex - startIndex);
                                    AppendOutput($"{Path.GetFileName(zipFile)} - 该压缩包文本的第 { lineIndex } 行包含“笔者”: {context}");
                                }
                                if (line.Contains("?"))// 检测文本中是否包含“?”
                                {
                                    int startIndex = Math.Max(0, line.IndexOf("?") - 8);
                                    int endIndex = Math.Min(line.Length, line.IndexOf("?") + 8 + 1);
                                    string context = line.Substring(startIndex, endIndex - startIndex);
                                    AppendOutput($"{Path.GetFileName(zipFile)} - 该压缩包文本的第 { lineIndex } 行包含“英文问号（?）”: {context}");
                                }
                            }

                            // 如果文本文件为空，输出提示信息
                            if (isEmpty)
                            {
                                AppendOutput($"{Path.GetFileName(zipFile)} - 该压缩包的 {entry.FullName} 文本文件内容为空");
                            }
                        }
                    }
                }
            }  
            // 查找重复的文本行
            foreach (KeyValuePair<string, List<int>> pair in lineIndexMap)
            {
                if (pair.Value.Count > 1)
                {
                    repeatedLines.Add(pair.Key);
                }
            }
            // 如果有重复的文本行，输出压缩文件名称和重复内容所在行数
            if (repeatedLines.Count > 0)
            {
                StringBuilder outputBuilder = new StringBuilder();
                outputBuilder.AppendLine($"{Path.GetFileName(zipFile)} - 压缩包内存在重复的文本行：");

                foreach (string line in repeatedLines)
                {
                    
                    foreach (int index in lineIndexMap[line])
                    {
                        outputBuilder.Append("第"+$"{index}"+"行");
                    }
                    
                }
                AppendOutput(outputBuilder.ToString());
            }
        }

        private void CheckTextFileNumber(string zipFile)//文本的数字与压缩包的数字
        {
            using (ZipArchive archive = ZipFile.OpenRead(zipFile))
            {
                ZipArchiveEntry textFileEntry = archive.Entries.FirstOrDefault(entry =>
                    entry.FullName.EndsWith(".txt", StringComparison.OrdinalIgnoreCase));

                if (textFileEntry != null)
                {
                    string textFileName = Path.GetFileNameWithoutExtension(textFileEntry.Name);
                    string zipFileName = Path.GetFileNameWithoutExtension(zipFile);

                    if (!textFileName.Any(char.IsDigit) || !zipFileName.Any(char.IsDigit))
                    {
                        AppendOutput($"{Path.GetFileName(zipFile)} - 压缩包的数字编号与内部文本不匹配");
                    }
                    else
                    {
                        string textFileNumber = new string(textFileName.Where(char.IsDigit).ToArray());
                        string zipFileNumber = new string(zipFileName.Where(char.IsDigit).ToArray());

                        if (textFileNumber != zipFileNumber)
                        {
                            AppendOutput($"{Path.GetFileName(zipFile)} - 压缩包的数字编号与内部文本不匹配");
                        }
                    }
                }
            }
        }

        private void CheckNestedZipFiles(string zipFile)//检测是否存在嵌套压缩包
        {
            using (ZipArchive archive = ZipFile.OpenRead(zipFile))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                    {
                        AppendOutput($"{Path.GetFileName(zipFile)} - 存在嵌套压缩包: {entry.Name}");
                    }
                }
            }
        }

        /*private void CheckImageTextSpacing(string zipFile)
        {
            using (ZipArchive archive = ZipFile.OpenRead(zipFile))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase))
                    {
                        using (var imageStream = entry.Open())
                        {
                            using (var image = Image.FromStream(imageStream))
                            {
                                var ocrEngine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default);
                                using (var page = ocrEngine.Process(image, PageSegMode.Auto))
                                {
                                    var text = page.GetText();
                                    var lines = text.Split('\n');

                                    var spacing = GetAverageSpacing(lines);

                                    if (!IsSpacingConsistent(lines, spacing))
                                    {
                                        AppendOutput($"{zipFile} - {entry.FullName} - 图片中的文字字间距不一致");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private float GetAverageSpacing(string[] lines)
        {
            int totalSpacing = 0;
            int count = 0;

            for (int i = 0; i < lines.Length - 1; i++)
            {
                int spacing = Math.Abs(lines[i + 1].Length - lines[i].Length);
                totalSpacing += spacing;
                count++;
            }

            if (count > 0)
            {
                return (float)totalSpacing / count;
            }
            else
            {
                return 0;
            }
        }

        private bool IsSpacingConsistent(string[] lines, float spacing)
        {
            for (int i = 0; i < lines.Length - 1; i++)
            {
                int currentSpacing = Math.Abs(lines[i + 1].Length - lines[i].Length);
                if (Math.Abs(currentSpacing - spacing) > 1) // 允许1个像素的误差
                {
                    return false;
                }
            }
            return true;
        }*/

        private void AppendOutput(string text)
        {
            outputBuilder.AppendLine(text);
            jisu++;
        }
        public string GetOutput()
        {
            return outputBuilder.ToString();
        }
    }
}