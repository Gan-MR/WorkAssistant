using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 工作助手.Function
{

    public class ScanXls
    {
        private StringBuilder outputBuilder;

        public ScanXls()
        {
            outputBuilder = new StringBuilder();
        }

        public static CancellationTokenSource cancellationTokenSource;
        public async Task CheckExcelLinks(string filePath, int startRow, int endRow, int targetColumn, bool highlightInvalidLinks, ProgressBar progressBar)
        {
            cancellationTokenSource = new CancellationTokenSource(); // 创建新的 CancellationTokenSource
            cancellationTokenSource.Token.ThrowIfCancellationRequested(); // 检查是否取消
            try
            {
                int jisu = 0;//统计打不开的链接
                int scanyes = 0;//统计已扫描的链接
                try
                {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                    {
                        XSSFWorkbook workbook = new XSSFWorkbook(fs);
                        XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(0); // 获取第一个工作表

                        SemaphoreSlim semaphore = new SemaphoreSlim(999); // 限制并行验证的线程数量为10

                        for (int i = startRow - 1; i < endRow; i++)
                        {
                            await semaphore.WaitAsync(); // 等待信号量
                            try
                            {


                                cancellationTokenSource.Token.ThrowIfCancellationRequested(); // 检查是否取消
                                await Task.Run(async () =>
                                {
                                    try
                                    {
                                        IRow row = sheet.GetRow(i);
                                        if (row != null)
                                        {
                                            string url = row.GetCell(targetColumn - 1)?.StringCellValue;
                                            if (!string.IsNullOrEmpty(url))
                                                scanyes++;
                                            {
                                                if (!await UrlIsValidAsync(url, 3, 10000, cancellationTokenSource.Token)) // 使用异步方式验证链接,参数URL，单链接扫描次数，超时时间（毫秒）
                                                {
                                                    lock (this) // 使用锁来保护共享资源 jisu
                                                    {
                                                        AppendOutput($"第 {i + 1} 行的链接 {url} 无法打开");
                                                        Interlocked.Increment(ref jisu);
                                                        if (highlightInvalidLinks)
                                                        {
                                                            HighlightRow(sheet, i);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    finally
                                    {
                                        semaphore.Release(); // 释放信号量
                                        progressBar.Invoke((MethodInvoker)delegate { progressBar.Value = i; });
                                    }
                                }, cancellationTokenSource.Token); // 传递 CancellationToken;
                            }
                            catch
                            {
                                if (cancellationTokenSource.Token.IsCancellationRequested) // 检查是否取消
                                {

                                    throw; // 如果取消操作，则跳出循环
                                }
                            }
                        }

                        await Task.Run(() => SaveExcelFile(workbook, filePath)); // 使用Task.Run在后台线程中执行保存操作
                    }
                }
                catch (Exception ex)
                {
                    // 处理异常
                    //AppendOutput($"发生错误: {ex.Message}");
                    if (ex.Message.Contains("另一进程") == true)
                    {
                        AppendOutput("所选择的表格正在被另一进程占用，扫描或染色可能不成功");
                    }
                }
                progressBar.Value = progressBar.Maximum;//拉满进度条
                if (jisu == 0) { AppendOutput("扫描已完成，本次扫描了" + scanyes + "个链接，没有找到打不开的链接"); }
                else { AppendOutput("扫描已完成，本次扫描了" + scanyes + "个链接，找到 " + jisu + " 条打不开的链接"); };

            }
            catch (OperationCanceledException)
            {
                // 处理取消操作
                progressBar.Value = progressBar.Maximum;//反向拉满进度条
            }

        }

        private void SaveExcelFile(XSSFWorkbook workbook, string filePath)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }

        private async Task<bool> UrlIsValidAsync(string url, int retryCount, int timeoutMilliseconds, CancellationToken cancellationToken)
        {
            for (int i = 0; i < retryCount; i++)
            {
                cancellationToken.ThrowIfCancellationRequested(); // 检查是否取消
                try
                {
                    using (var client = new HttpClient())
                    {
                        client.Timeout = TimeSpan.FromMilliseconds(timeoutMilliseconds);
                        var response = await client.GetAsync(url);

                        if (response.IsSuccessStatusCode)
                        {
                            return true;
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    throw; // 取消操作时抛出 OperationCanceledException
                }
                catch (Exception)
                {
                    // 处理或记录异常
                }
            }

            return false; // 如果重试后仍无法打开，则判定为无效链接
        }

        //表格染色Start-----------------------------
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
        //表格染色End---------------------------------

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

