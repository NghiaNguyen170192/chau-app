using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Drawing;
using System.IO;
using System.Net;

namespace ChauApp
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var input = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data");
            //var output = @"C:\Users\DELL\Desktop";
            var output = @"E:\";
            var files = Directory.GetFiles(input, "*.xlsx");

            foreach (var file in files)
            {
                var workBook = ProcessWorkBook(file);
                var fileName = Path.GetFileName(file);
                WriteToFile(output, fileName, workBook);
            }
        }

        private static IWorkbook ProcessWorkBook(string file)
        {
            var workBook = GetWorkbook(file);
            ISheet sheet = workBook.GetSheetAt(0);

            for (var rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null) continue;

                var firstCell = row.GetCell(0);
                if (firstCell == null || firstCell.StringCellValue == null) continue;
                var url = GetUrl(firstCell.StringCellValue);
                if (string.IsNullOrEmpty(url)) continue;

                for (var cellIndex = 0; cellIndex <= row.LastCellNum; cellIndex++)
                {
                    var cell = row.GetCell(cellIndex);
                    if (cell == null || cell.StringCellValue == null) continue;

                    if (cell.StringCellValue.Contains(url) || !IsImageUrlValid(url))
                    {
                        cell.SetCellValue(cell.StringCellValue.Replace(url, ""));
                    }
                }
            }

            return workBook;
        }

        private static bool IsImageUrlValid(string url)
        {
            try
            {
                using (var webClient = new WebClient())
                {
                    var imageData = webClient.DownloadData(url);
                    if (imageData.Length <= 0) return false;

                    using (var imgStream = new MemoryStream(imageData))
                    {
                        var img = Image.FromStream(imgStream);
                        if (img.Width >= 350 && img.Height >= 350)
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{url}] - {ex.Message}");
            }

            return false;
        }

        private static void WriteToFile(string output, string fileName, IWorkbook workBook)
        {
            fileName = fileName.Replace(".xlsx", "-output.xlsx");
            var path = Path.Combine(output, fileName);

            using (FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                workBook.Write(fileStream);
            }
        }

        private static string GetUrl(string text)
        {
            var url = string.Empty;
            if (string.IsNullOrEmpty(text)) return url;

            if (!text.Contains("https")) return url;

            url = text.Trim().Substring(text.IndexOf("https"));

            if (url.Contains("-IMAGE"))
            {
                url = url.Substring(0, url.IndexOf("-IMAGE"));
            }

            if (url.Contains("-Media"))
            {
                url = url.Substring(0, url.IndexOf("-Media"));
            }

            if (url.Contains("-HTTP"))
            {
                url = url.Substring(0, url.IndexOf("-HTTP"));
            }

            return url;
        }

        private static IWorkbook GetWorkbook(string file)
        {
            IWorkbook workbook;
            using (FileStream fileStream = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fileStream);
            }

            return workbook;
        }
    }
}

