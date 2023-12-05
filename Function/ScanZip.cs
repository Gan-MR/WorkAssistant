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

        public void ScanDirectory(string directoryPath, bool CheckTxtData)
        {
            try
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

                    CheckTextFileCountAndImageCount(zipFile);//图片数量>4，文本=1，且文本行数=图片张数，文本第一行少于四个字，文本命名规范【XXX】本题图片文本，文本文件编号与压缩包编号匹配
                    if (!CheckTxtData)
                        CheckTextFileContent(zipFile);//压缩包内文本超过17个连续字符相同字符，文本是否为空，是否包含空格，笔者，小编，英文问号等关键字
                    CheckNestedZipFiles(zipFile);//嵌套压缩包
                    zipjisu++;
                }
                if (GetOutput().Length == 0)
                {

                    AppendOutput("扫描已完成，扫描了" + zipjisu + "个压缩包，没有发现已知问题");

                }
                else
                {
                    AppendOutput("本次扫描了 " + zipjisu + " 个压缩包，发现了 " + bugjisu + " 个问题");
                }
            }
            catch (Exception) { AppendOutput("扫描启动失败，未设定目录或选择所目录异常"); }
        }

        public void CheckTextFileCountAndImageCount(string zipFilePath)
        {
            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
                {
                    // 检查压缩包内是否只包含一个文本文件
                    ZipArchiveEntry textFileEntry = null;
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        if (entry.FullName.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                        {
                            if (textFileEntry != null)
                            {
                                AppendOutput(Path.GetFileName(zipFilePath) + " - 包含多个文本文件。");

                            }
                            textFileEntry = entry;
                        }
                    }
                    if (textFileEntry == null)
                    {
                        AppendOutput(Path.GetFileName(zipFilePath) + " - 没有找到文本文件。");
                        //return;
                    }

                    // 检查文本文件的行数是否与图片文件的数量匹配
                    int imageCount = 0;
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        if (!entry.FullName.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                        {
                            imageCount++;
                        }
                    }
                    using (Stream stream = textFileEntry.Open())
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        int lineCount = 0;
                        while (!reader.EndOfStream)
                        {
                            reader.ReadLine();
                            lineCount++;
                        }
                        if (lineCount != imageCount)
                        {
                            AppendOutput(Path.GetFileName(zipFilePath) + " - 图片与文本行数不符，发现" + imageCount + "个.jpg图片文件，" + lineCount + "行文本。");
                            //return;
                        }
                    }

                    // 检查压缩包内是否至少包含四个图片文件
                    if (imageCount < 4)
                    {
                        AppendOutput(Path.GetFileName(zipFilePath) + " - 包含少于4个图片文件。");
                        //return;
                    }

                    // 检查文本文件的第一行内容是否至少包含4个字符
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        using (StreamReader reader = new StreamReader(entry.Open(), Encoding.Default))
                        {
                            string firstLine = reader.ReadLine();
                            if (firstLine != null && firstLine.Length < 4)
                            {
                                AppendOutput($"{Path.GetFileName(zipFilePath)} - .txt文本的第一行小于4个字");
                            }


                        }
                    }

                    // 检查压缩包的文件名是否与文本文件名匹配
                    string archiveName = Path.GetFileNameWithoutExtension(zipFilePath);
                    string textFileName = Path.GetFileNameWithoutExtension(textFileEntry.FullName);
                    int bracketStart = textFileName.LastIndexOf('【');
                    int bracketEnd = textFileName.LastIndexOf('】');
                    if (bracketStart >= 0 && bracketEnd > bracketStart)
                    {
                        Match match = Regex.Match(archiveName, @"(.+?)\p{IsCJKUnifiedIdeographs}+", RegexOptions.IgnoreCase);
                        string expectedName = textFileName.Substring(bracketStart + 1, bracketEnd - bracketStart - 1);
                        if (!match.Groups[1].Value.EndsWith(expectedName))
                        {
                            AppendOutput(Path.GetFileName(zipFilePath) + " - 与压缩包内的文本命名 " + expectedName + " 不匹配。");
                            return;
                        }
                    }
                    if (Regex.IsMatch(textFileName, @"^【.*】本题图片文本$"))
                    {

                    }
                    else
                    {
                        AppendOutput(Path.GetFileName(zipFilePath) + " - 文本文件不符合“【XXX】本题图片文本”命名要求");
                    }
                    //AppendOutput(Path.GetFileName(zipFilePath)+" - 通过所有检查。");
                }
            }
            catch (Exception ex)
            {
                AppendOutput(string.Format("扫描压缩包 " + Path.GetFileName(zipFilePath) + " 时出现错误：" + ex.Message + "\r\n可能该压缩包没有符合规格的文件或者没有授予访问权限"));
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
                                CheckAndAppendOutput(zipFile, lineIndex, line, " ", "空格");
                                CheckAndAppendOutput(zipFile, lineIndex, line, "小编", "小编");
                                CheckAndAppendOutput(zipFile, lineIndex, line, "笔者", "笔者");
                                CheckAndAppendOutput(zipFile, lineIndex, line, "?", "英文问号（?）");
                            }

                            // 如果文本文件为空，输出提示信息
                            if (isEmpty)
                            {
                                AppendOutput($"{Path.GetFileName(zipFile)} - 该压缩包的文本文件内容为空");
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
                        outputBuilder.Append("第" + $"{index}" + "行");
                    }

                }
                AppendOutput(outputBuilder.ToString());
            }
        }

        private void CheckAndAppendOutput(string zipFile, int lineIndex, string line, string keyword, string message)//外置查找非法内容
        {
            if (line.Contains(keyword))
            {
                int startIndex = Math.Max(0, line.IndexOf(keyword) - 8);
                int endIndex = Math.Min(line.Length, line.IndexOf(keyword) + 8 + 1);
                string context = line.Substring(startIndex, endIndex - startIndex);
                AppendOutput($"{Path.GetFileName(zipFile)} - 该压缩包文本的第 {lineIndex} 行包含“{message}”: {context}");
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
            bugjisu++;
        }
        public string GetOutput()
        {
            return outputBuilder.ToString();
        }
    }
}