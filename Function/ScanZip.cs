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

        public void ScanDirectory(string directoryPath, string keyWord)
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
                    if (keyWord != "")
                        CheckTextFileContent(zipFile, keyWord);//压缩包内文本超过17个连续字符相同字符，文本是否为空，是否包含关键字
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
                        return;
                    }

                    // 检查文本文件的行数是否与图片文件的数量匹配
                    int imageCount = 0;
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        if (entry.FullName.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase))
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
                        if (lineCount < imageCount)
                        {
                            AppendOutput(Path.GetFileName(zipFilePath) + " - 图片与文本行数不符，发现" + imageCount + "个.jpg图片文件，" + lineCount + "行文本。");
                            //return;
                        }
                    }

                    // 检查压缩包内是否至少包含四个图片文件
                    if (imageCount < 4)
                    {
                        AppendOutput(Path.GetFileName(zipFilePath) + " - 所包含的.jpg图片少于4个。");
                        //return;
                    }

                    /*
                     //检查文本文件的第一行内容是否至少包含4个字符
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
                    }*/

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
            catch (Exception)
            {
                AppendOutput("扫描压缩包 " + Path.GetFileName(zipFilePath) + " 时出现错误：可能该压缩包没有符合规格的文件或者没有授予访问权限");
            }


        }

        private void CheckNestedZipFiles(string zipFile)//检测是否存在嵌套压缩包
        {
            bool nestedZipFound = false; // 用于标记是否找到嵌套压缩包

            // 检测压缩包文件名是否包含空格
            if (Path.GetFileName(zipFile).Contains(" "))
            {
                AppendOutput($"{Path.GetFileName(zipFile)} - 压缩包文件名包含空格");
            }

            using (ZipArchive archive = ZipFile.OpenRead(zipFile))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                    {
                        nestedZipFound = true;
                        break; // 如果找到嵌套压缩包，立即退出循环
                    }
                }
            }

            if (nestedZipFound)
            {
                AppendOutput($"{Path.GetFileName(zipFile)} - 存在嵌套压缩包");
            }
        }

        private void CheckTextFileContent(string zipFile, string Keyword)
        /* 压缩包内文本超过17个连续字符相同字符，文本是否为空，是否包含空格
         * 文本行最后一个字符必须为。！”？
         * 引号成对。
         */
        {
            List<string> repeatedLines = new List<string>(); // 用于存储重复的文本行
            Dictionary<string, List<int>> lineIndexMap = new Dictionary<string, List<int>>(); // 用于存储每个文本行的出现位置
            string[] keyword = Keyword.Split('.');

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
                            int lineNumber = 0;
                            bool isEmpty = true; // 用于标记文本文件是否为空
                            string lastLine = null; // 用于存储最后一行的内容
                            while ((line = reader.ReadLine()) != null)
                            {
                                lineNumber++;
                                lineIndex++;
                                isEmpty = false; // 当读取到文本行时，将 isEmpty 标记为 false

                                // 检测引号成对
                                if (ContainsUnmatchedQuote(line))
                                {
                                    AppendOutput($"{Path.GetFileName(zipFile)} - 第 {lineNumber} 行，引号不成对");
                                }

                                //先判断文字> 17
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
                                lastLine = line; // 每次读取新的一行时，更新最后一行的内容

                                // 检测文本中是否包含。。。。
                                //CheckAndAppendOutput(压缩包, 所在行, line, "关键字");
                                for (int i = 0; i < keyword.Length; i++)
                                {
                                    CheckAndAppendOutput(zipFile, lineIndex, line, keyword[i]);
                                }

                                /*
                                 * 检测文本行(除第一行)的最后一个字是否为！？。”
                                 * 最后一个字不是中文下引号，那么倒数第二个字则不能是句号，感叹号，问号
                                 */
                                if (lineNumber > 1 && line.Length >= 9) // 从第二行开始检测并且只检测文本行长度大于等于9个字符的行
                                {
                                    if (!string.IsNullOrEmpty(line))
                                    {
                                        char lastChar = line[line.Length - 1];
                                        if (lastChar != '”')
                                        {
                                            char secondLastChar = line.Length > 1 ? line[line.Length - 2] : '\0';
                                            if (secondLastChar == '。' || secondLastChar == '！' || secondLastChar == '？')
                                            {
                                                AppendOutput($"{Path.GetFileName(zipFile)} - 第 {lineNumber} 行，倒数第二个字符为：{secondLastChar}");
                                            }
                                        }
                                        else
                                        {
                                            char secondLastChar = line.Length > 1 ? line[line.Length - 2] : '\0';
                                            string surroundingChars = line.Substring(line.Length - 9, 9);
                                            if (secondLastChar != '。' && secondLastChar != '！' && secondLastChar != '？')
                                            {
                                                AppendOutput($"{Path.GetFileName(zipFile)} - 第 {lineNumber} 行，下引号前没有合理的结束：{surroundingChars}");
                                            }
                                        }
                                    }

                                    if (!string.IsNullOrEmpty(line))
                                    {
                                        char lastChar = line[line.Length - 1];
                                        if (lastChar == '！' || lastChar == '。' || lastChar == '”' || lastChar == '？')
                                        {
                                            // 文本行的最后一个字符符合条件
                                        }
                                        else
                                        {
                                            AppendOutput($"{Path.GetFileName(zipFile)} - 第 {lineNumber} 行，最后一个字符为：{lastChar}");
                                        }
                                    }
                                }
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
                outputBuilder.Append($"{Path.GetFileName(zipFile)} - 存在重复的文本行：");

                foreach (string line in repeatedLines)
                {

                    foreach (int index in lineIndexMap[line])
                    {
                        outputBuilder.Append("第 " + $"{index}" + " 行 ");
                    }

                }
                AppendOutput(outputBuilder.ToString());
            }
        }
        // 检测引号成对的方法
        private bool ContainsUnmatchedQuote(string line)
        {
            int countOpen = 0;
            int countClose = 0;
            foreach (char c in line)
            {
                if (c == '“')
                {
                    countOpen++;
                }
                else if (c == '”')
                {
                    countClose++;
                }
            }
            return countOpen != countClose; // 如果开引号数量不等于闭引号数量，则返回 true，表示引号不成对
        }
        //外置查询关键字
        private void CheckAndAppendOutput(string zipFile, int lineIndex, string line, string keyword)
        {
            Regex regex = new Regex(Regex.Escape(keyword) + "+");
            Match match = regex.Match(line);
            if (match.Success)
            {
                string matchedKeyword = match.Value;
                int index = 0;
                while ((index = line.IndexOf(matchedKeyword, index)) != -1)
                {
                    int startIndex = Math.Max(0, index - 11);
                    int endIndex = Math.Min(line.Length, index + matchedKeyword.Length + 11);
                    string context = line.Substring(startIndex, endIndex - startIndex);
                    AppendOutput($"{Path.GetFileName(zipFile)} - 第 {lineIndex} 行包含关键字：》{matchedKeyword}《；位置: {context}");
                    index += matchedKeyword.Length;
                }
            }
            else
            {
                int index = 0;
                while ((index = line.IndexOf(keyword, index)) != -1)
                {
                    string matchedKeyword = line.Substring(index, keyword.Length);
                    int startIndex = Math.Max(0, index - 11);
                    int endIndex = Math.Min(line.Length, index + keyword.Length + 11);
                    string context = line.Substring(startIndex, endIndex - startIndex);
                    AppendOutput($"{Path.GetFileName(zipFile)} - 第 {lineIndex} 行包含关键字：》{matchedKeyword}《；位置: {context}");
                    index += keyword.Length;
                }
            }
        }




    }
}