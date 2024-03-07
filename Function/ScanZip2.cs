using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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
        public void CheckZip(string dir1, string dir2)
        {
           
                if (Directory.Exists(dir1))
                {
                    CheckDirectoryNumbers(dir1);
                }
                else
                {
                    AppendOutput("目录1：找不到对象");
                }

            
           

        }

        public void CheckDirectoryNumbers(string directoryPath)
        {
            // 获取目录下所有文件夹
            string[] directories = Directory.GetDirectories(directoryPath);

            // 提取所有文件夹名称中的数字部分
            List<int> numbers = new List<int>();
            foreach (string directory in directories)
            {
                string folderName = new DirectoryInfo(directory).Name;
                Match match = Regex.Match(folderName, @"\d+");
                if (match.Success)
                {
                    int number = int.Parse(match.Value);
                    numbers.Add(number);
                }
            }

            // 校验数字是否连续
            numbers.Sort();
            int missingCount = 0;
            bool hasException = true;
            for (int i = 1; i < numbers.Count; i++)
            {
                if (numbers[i] != numbers[i - 1] + 1)
                {
                    // 输出相邻的两个文件夹名
                    int prevNumber = numbers[i - 1];
                    int nextNumber = numbers[i];
                    string prevDirectory = directories.First(d => new DirectoryInfo(d).Name.Contains(prevNumber.ToString()));
                    string nextDirectory = directories.First(d => new DirectoryInfo(d).Name.Contains(nextNumber.ToString()));

                    int count = numbers[i] - numbers[i - 1] - 1;
                    missingCount += count;
                    if (count > 100)
                    {
                        AppendOutput("存在区间异常，可能存在附加的任务或编号错误");
                        hasException = false;
                    }
                    else
                    {
                        AppendOutput("在 " + new DirectoryInfo(prevDirectory).Name + " 和 " + new DirectoryInfo(nextDirectory).Name + " 之间，缺少 " + count + " 个文件夹");
                    }
                }
            }

            if (hasException)
            {
                if (missingCount != 0)
                {
                    AppendOutput("总共缺少 " + missingCount + " 个文件夹。");
                }
                else
                {
                    AppendOutput("没有找到不连续的文件夹。");
                }
            }
        }


        public void CompareDataWithFolders(string xlsPath, string specifiedColumn, List<string> folderDirectories)
        {
            int totalFoldersFound = 0; // 扫描总量
            int missingCount = 0; // 缺失量
            List<string> columnData = new List<string>();
            string xlsLine;

            // 打开用户输入的xls表格
            using (FileStream file = new FileStream(xlsPath, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(file);
                ISheet sheet = workbook.GetSheetAt(0);

                // 获取用户指定列的所有数据
                int columnIndex = specifiedColumn[0] - 'A';

                for (int i = 0; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null && row.GetCell(columnIndex) != null)
                    {
                        ICell cell = row.GetCell(columnIndex);
                        switch (cell.CellType)
                        {
                            case CellType.Numeric:
                                columnData.Add(cell.NumericCellValue.ToString());
                                break;
                            case CellType.String:
                                columnData.Add(cell.StringCellValue);
                                break;
                            // 可根据需要添加其他类型的处理
                            default:
                                columnData.Add(cell.ToString());
                                break;
                        }
                    }
                }

                List<string> allFolderNames = new List<string>();
                bool hasValidDirectory = false; // 有效目录标识
                xlsLine = columnData.Count.ToString();

                // 遍历文件夹目录
                foreach (string directory in folderDirectories)
                {
                    if (Directory.Exists(directory))
                    {
                        string[] folders = Directory.GetDirectories(directory);
                        allFolderNames.AddRange(folders.Select(Path.GetFileName));
                        hasValidDirectory = true;
                    }
                    else
                    {
                        // 输出文件夹目录不存在
                        AppendOutput($"目录 “{directory}” :找不到这个目录");
                    }
                }

                if (!hasValidDirectory) return; // 如果没有一个有效目录就不继续

                totalFoldersFound = allFolderNames.Count;
                foreach (string data in columnData)
                {
                    bool foundInFolders = false;
                    foreach (string folderName in allFolderNames)
                    {
                        if (folderName.Contains(data) || data.Contains(folderName))
                        {
                            foundInFolders = true;
                            break;
                        }
                    }
                    if (!foundInFolders)
                    {
                        // 输出对应的数据
                        AppendOutput($"表格中的数据 {data} 没有匹配的文件夹名");
                        missingCount++;
                    }
                }
            }

            int xlsLineCount = int.Parse(xlsLine);
            AppendOutput($"所选表格A列有 {xlsLine} 行数据。检测到 {totalFoldersFound} 个文件夹，匹配了 {xlsLineCount - missingCount} 个文件夹。缺失 {missingCount} 个文件夹");
        }











    }
}


