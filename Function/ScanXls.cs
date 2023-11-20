using System;
using System.IO;
using System.Text;
using System.Net;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;

namespace 工作助手.Function
{

    public class ScanXls
    {
        private StringBuilder outputBuilder;

        public ScanXls()
        {
            outputBuilder = new StringBuilder();
        }


        public void CheckExcelLinks(string filePath, int startRow, int endRow, int targetColumn, bool highlightInvalidLinks)
        {
            int yueorno = 0;
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    XSSFWorkbook workbook = new XSSFWorkbook(fs);
                    XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(0); // 获取第一个工作表

                    for (int i = startRow - 1; i < endRow && i <= sheet.LastRowNum; i++)
                    {
                        
                        IRow row = sheet.GetRow(i);
                        if (row != null)
                        {
                            string url = row.GetCell(targetColumn - 1)?.StringCellValue;
                            if (!string.IsNullOrEmpty(url))
                            {
                                if (!UrlIsValid(url))
                                {
                                    AppendOutput($"第 {i + 1} 行的链接 {url} 无法打开");
                                    yueorno++;
                                    if (highlightInvalidLinks)
                                    {
                                        HighlightRow(sheet, i);
                                    }
                                }
                            }
                        }
                    }
                    SaveExcelFile(workbook, filePath); // 保存对文件的更改
                }
            }
            catch (Exception ex)
            {
                AppendOutput($"发生错误: {ex.Message}");
            }
            if (yueorno == 0) { AppendOutput("没有找到打不开的链接"); }
            else { AppendOutput("总共找到 " + yueorno + " 条打不开的链接"); };
        }

        private void SaveExcelFile(XSSFWorkbook workbook, string filePath)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                workbook.Write(memoryStream);
                File.WriteAllBytes(filePath, memoryStream.ToArray());
            }
        }
        private bool UrlIsValid(string url)
        {
            // 使用正则表达式验证URL是否为网站地址
            string pattern = @"^(https?|ftp)://[^\s/$.?#].[^\s]*$";
            if (Regex.IsMatch(url, pattern))
            {
                try
                {
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                    request.Method = "HEAD"; // 使用HEAD方法，只获取响应头，不下载页面内容
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        return response.StatusCode == HttpStatusCode.OK;
                    }
                }
                catch
                {
                    return false;
                }
            }
            else
            {
                AppendOutput($"目标 {url} 不是一个网站链接");
                return true;
            }
        }

        private void HighlightRow(XSSFSheet sheet, int rowIndex)
        {
            XSSFRow row = (XSSFRow)sheet.GetRow(rowIndex);
            if (row != null)
            {
                XSSFCellStyle style = (XSSFCellStyle)sheet.Workbook.CreateCellStyle();
                style.FillForegroundColor = IndexedColors.Yellow.Index;
                style.FillPattern = FillPattern.SolidForeground;

                for (int i = 0; i < row.LastCellNum; i++)
                {
                    XSSFCell cell = (XSSFCell)row.GetCell(i);
                    if (cell != null)
                    {
                        cell.CellStyle = style;
                    }
                }
            }
        }





        private void AppendOutput(string text)
        {
            outputBuilder.AppendLine(text);
        }

        public string GetOutput()
        {
            return outputBuilder.ToString();
        }



    }
}

