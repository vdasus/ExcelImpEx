using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ExcImp
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var rez = GetNonEmptyDistinctData<DataItem>(@"D:\REPO\_Learn\ExcImp\InputData.xlsx", false, new Dictionary<string, string>()
                {
                    {"įmonės kodas", "Code"},
                    {"TP valst. Nr", "Number"},
                    {"TP VIN kodas", "Vin"}
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        public static IReadOnlyCollection<T> GetData<T>(string fileName, bool throwOnFirstError = false, Dictionary<string, string> map = null) where T : new()
        {
            FileInfo existingFile = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                return worksheet.ToList<T>(throwOnFirstError, map).AsReadOnly();
            }
        }

        public static IReadOnlyCollection<T> GetNonEmptyDistinctData<T>(string fileName, bool throwOnFirstError = false, Dictionary<string, string> map = null) where T : new()
        {
            return GetData<T>(fileName, throwOnFirstError, map)
                .Where(s => !Extentions.CheckIsObjectEmpty(s))
                .Distinct()
                .ToList();
        }

        public static IReadOnlyCollection<T> GetNonEmptyData<T>(string fileName, bool throwOnFirstError = false, Dictionary<string, string> map = null) where T : new()
        {
            return GetData<T>(fileName, throwOnFirstError, map)
                .Where(s => !Extentions.CheckIsObjectEmpty(s))
                .ToList();
        }

    }
}
