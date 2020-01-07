using Oracle.ManagedDataAccess.Client;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace XuatBaoCao
{
    class Program
    {
        const string SQL_TEXT_REGEX = "SQL\\((.*?)\\)$";
        static readonly Regex regex = new Regex(SQL_TEXT_REGEX);
        // const string connStr = "User Id=LOS_UAT_2;Password=admin123;Data Source=devdb.ists.com.vn:1523/oracledb";
        const string connStr = "User Id=LOS_UAT3;Password=losuat3abcd1234;Data Source=devdb.ists.com.vn:1522/devdb";
        static void Main(string[] args)
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("MTg3MTE5QDMxMzcyZTM0MmUzMGJQVEVTWEVxVHNXMmlKdVhhWWZuSGg5UUtWOTFaa2YvSEJzSGJURVFyeUE9;MTg3MTIwQDMxMzcyZTM0MmUzMGZtNEJvZW81cFZoUUVJZHhCUUFWeFUrbVZWZ2s5anlVVzNRaGRZUHBNUk09;MTg3MTIxQDMxMzcyZTM0MmUzMGVPbnI1WHlpL1hTSHg5MWFXZld2ZC91VVhDNUUrWjluVEtlSm1lU0U1NTQ9;MTg3MTIyQDMxMzcyZTM0MmUzMFNrSmhCem5uOTBHcWNQT0UwWU93TE9GZzd4Ykc0eGJ4M1llYkxEYldDKzA9;MTg3MTIzQDMxMzcyZTM0MmUzMFUyNUh2S0tRK1BEOHFYZDhyQllncnhRbE9GWC9jWjJuSDZEZ1RGNjRuMTg9;MTg3MTI0QDMxMzcyZTM0MmUzMGllZDZvVnhwUTRUQmV3Ry9BSnd3NWdFaTZaVTRwZGpxZmNINVRjak0vQ009;MTg3MTI1QDMxMzcyZTM0MmUzMFRSNUYrazU1Wkd0OUl3bkhkRXRpOWxKbzEzWjNERDhKeTMvb21HaXlyeXM9;MTg3MTI2QDMxMzcyZTM0MmUzMHBEdEl5WUlkdEVqamtZajBWd0RMYlZSdk1rTFROMWp5ZS94RDdzNThoc1k9;MTg3MTI3QDMxMzcyZTM0MmUzME5NOG16WWdFOTRFUlEzVWdWcW9QWGNtVTFINDlUNmZ0MWVaZ1hHM0RvSFk9;NT8mJyc2IWhiZH1nfWN9YGpoYmF8YGJ8ampqanNiYmlmamlmanMDHmgxJjolMj19Izs8PTRqZxM0PjI6P30wPD4=;MTg3MTI4QDMxMzcyZTM0MmUzMEVTRlJYeWdMcHZndi92disvYXArVUwwbnd5R1ZwQmFnRmM3eG9Oc1BNMnc9");
            using (OracleConnection conn = new OracleConnection(connStr))
            {
                conn.Open();
                DoUpdateLog(conn);
                DoExportReport(conn, "Đồng bộ dư nợ.xlsx");
                conn.Close();
            }
        }
        public static void DoExportReport(OracleConnection conn, string filePath)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                IWorkbook workbook = application.Workbooks.Open(filePath);
                foreach (WorksheetImpl worksheet in workbook.Worksheets)
                {
                    DoFillDataIntoWorksheet(worksheet, conn);
                }
                //WorksheetImpl worksheet = (WorksheetImpl)workbook.Worksheets[0];


                workbook.SaveAs($"{Path.GetFileNameWithoutExtension(filePath)}_{DateTime.Now.ToString("yyyyMMdd")}.xlsx");
            }
        }
        private static void DoFillDataIntoWorksheet(WorksheetImpl worksheet, OracleConnection conn)
        {
            try
            {
                using (var cmd = new OracleCommand(worksheet.Range[$"A1"].Text, conn))
                {
                    DataTable table = new DataTable();
                    OracleDataAdapter da = new OracleDataAdapter(cmd);
                    da.Fill(table);
                    DataView view = table.DefaultView;
                    worksheet.ImportDataView(view, true, 1, 1);
                }
            }
            catch
            {

            }

        }



        public static void DoUpdateLog(OracleConnection conn)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                IWorkbook workbook = application.Workbooks.Open("ABBank_LOS_Tổng hợp kết quả import.xlsx");
                WorksheetImpl worksheet = (WorksheetImpl)workbook.Worksheets[0];
                for (int i = 1; i <= worksheet.LastRow; i++)
                {
                    FilterValue(conn, worksheet.Range[$"D{i}"]);
                }
                for (int i = 1; i <= worksheet.LastRow; i++)
                {
                    FilterValue(conn, worksheet.Range[$"E{i}"]);
                }

                ////Shifts cells towards Left after deletion
                //worksheet.Range["A1:E1"].Clear(ExcelMoveDirection.MoveLeft);
                ////Shifts cells toward Up after deletion
                //worksheet.Range["A1:A6"].Clear(ExcelMoveDirection.MoveUp);
                workbook.SaveAs($"ABBank_LOS_Tổng hợp kết quả import_{DateTime.Now.ToString("yyyyMMdd")}.xlsx");
                //Process.Start("EXCEL.EXE", "Book1.xlsx");
            }
        }
        public static void FilterValue(OracleConnection conn, IRange range)
        {

            string content = range.Text;
            if (!string.IsNullOrEmpty(content) && regex.IsMatch(content))
            {
                string sql = regex.Match(content).Groups[1].Value;
                using (var cmd = new OracleCommand(sql, conn))
                {
                    object val = cmd.ExecuteScalar();
                    double convertedVal;
                    if (double.TryParse(val.ToString(), out convertedVal))
                        range.Number = Convert.ToDouble(cmd.ExecuteScalar());
                    else range.Text = val.ToString();
                }
            }

        }

    }

    public static class ExcelUtil
    {
        public static void SaveAs(this IWorkbook workbook, string fileName)
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
            FileStream f = File.Open(fileName, FileMode.CreateNew);
            workbook.SaveAs(f);
            workbook.Close();
            f.Close();
        }
        public static IWorkbook Open(this IWorkbooks workbook, string fileName)
        {
            FileStream f = File.Open(fileName, FileMode.Open);
            return workbook.Open(f);
        }
    }
}
