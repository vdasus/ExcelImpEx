using System.ComponentModel.DataAnnotations;
using System;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace gMotoImpEx.Core.DomainServices
{
/*public class DataItem
        {
            public string Code { get; set; }
            public string Number { get; set; }
            public string Vin { get; set; }
        }

    var rez = tmp.GetData<DataItem>(@"D:\REPO\_Learn\ExcImp\InputData.xlsx", new Dictionary<string, string>()
    {
    {"įmonės kodas", "Code"},
    {"TP valst. Nr", "Number"},
    {"TP VIN kodas", "Vin"}
    });*/

    public class ExcelDataSource
    {
        #region Implementation of IDataSource

        public IReadOnlyCollection<T> GetData<T>(string fileName, Dictionary<string, string> map = null) where T : new()
        {
            FileInfo existingFile = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                return worksheet.ToList<T>(map).AsReadOnly();
            }
        }
        #endregion
    }

    public class ExcelMap
    {
        public string Name { get; set; }
        public string MappedTo { get; set; }
        public int Index { get; set; }
    }

    public static class Extentions
    {
        public static List<T> ToList<T>(this ExcelWorksheet worksheet, Dictionary<string, string> map = null) where T : new()
        {
            //DateTime Conversion
            var convertDateTime = new Func<double, DateTime>(excelDate =>
            {
                if (excelDate < 1)
                    throw new ArgumentException("Excel dates cannot be smaller than 0.");

                var dateOfReference = new DateTime(1900, 1, 1);

                if (excelDate > 60d)
                    excelDate = excelDate - 2;
                else
                    excelDate = excelDate - 1;
                return dateOfReference.AddDays(excelDate);
            });

            var props = typeof(T).GetProperties()
                .Select(prop =>
                {
                    var displayAttribute = (DisplayAttribute)prop.GetCustomAttributes(typeof(DisplayAttribute), false).FirstOrDefault();
                    return new
                    {
                        prop.Name,
                        DisplayName = displayAttribute == null ? prop.Name : displayAttribute.Name,
                        Order = displayAttribute?.GetOrder() == null ? 999 : displayAttribute.Order,
                        PropertyInfo = prop,
                        prop.PropertyType,
                        HasDisplayName = displayAttribute != null
                    };
                })
            .Where(prop => !string.IsNullOrWhiteSpace(prop.DisplayName))
            .ToList();

            var retList = new List<T>();
            var columns = new List<ExcelMap>();

            var start = worksheet.Dimension.Start;
            var end = worksheet.Dimension.End;
            var startCol = start.Column;
            var startRow = start.Row;
            var endCol = end.Column;
            var endRow = end.Row;

            // Assume first row has column names
            for (int col = startCol; col <= endCol; col++)
            {
                var cellValue = (worksheet.Cells[startRow, col].Value ?? string.Empty).ToString().Trim();
                if (!string.IsNullOrWhiteSpace(cellValue))
                {
                    columns.Add(new ExcelMap()
                    {
                        Name = cellValue,
                        MappedTo = map == null || map.Count == 0 ?
                            cellValue :
                            map.ContainsKey(cellValue) ? map[cellValue] : string.Empty,
                        Index = col
                    });
                }
            }

            // Now iterate over all the rows
            for (int rowIndex = startRow + 1; rowIndex <= endRow; rowIndex++)
            {
                var item = new T();
                columns.ForEach(column =>
                {
                    var value = worksheet.Cells[rowIndex, column.Index].Value;
                    var valueStr = value == null ? string.Empty : value.ToString().Trim();
                    var prop = string.IsNullOrWhiteSpace(column.MappedTo) ?
                        null :
                        props.First(p => p.Name.Contains(column.MappedTo));

                    // Excel stores all numbers as doubles, but we're relying on the object's property types
                    if (prop == null) return;

                    var propertyType = prop.PropertyType;
                    object parsedValue = null;

                    if (propertyType == typeof(int?) || propertyType == typeof(int))
                    {
                        if (!int.TryParse(valueStr, out var val))
                        {
                            val = default(int);
                        }

                        parsedValue = val;
                    }
                    else if (propertyType == typeof(short?) || propertyType == typeof(short))
                    {
                        if (!short.TryParse(valueStr, out var val))
                            val = default(short);
                        parsedValue = val;
                    }
                    else if (propertyType == typeof(long?) || propertyType == typeof(long))
                    {
                        if (!long.TryParse(valueStr, out var val))
                            val = default(long);
                        parsedValue = val;
                    }
                    else if (propertyType == typeof(decimal?) || propertyType == typeof(decimal))
                    {
                        if (!decimal.TryParse(valueStr, out var val))
                            val = default(decimal);
                        parsedValue = val;
                    }
                    else if (propertyType == typeof(double?) || propertyType == typeof(double))
                    {
                        if (!double.TryParse(valueStr, out var val))
                            val = default(double);
                        parsedValue = val;
                    }
                    else if (propertyType == typeof(DateTime?) || propertyType == typeof(DateTime))
                    {
                        if (value != null) parsedValue = convertDateTime((double)value);
                    }
                    else if (propertyType.IsEnum)
                    {
                        try
                        {
                            parsedValue = Enum.ToObject(propertyType, int.Parse(valueStr));
                        }
                        catch
                        {
                            parsedValue = Enum.ToObject(propertyType, 0);
                        }
                    }
                    else if (propertyType == typeof(string))
                    {
                        parsedValue = valueStr;
                    }
                    else
                    {
                        try
                        {
                            parsedValue = Convert.ChangeType(value, propertyType);
                        }
                        catch
                        {
                            parsedValue = valueStr;
                        }
                    }

                    try
                    {
                        prop.PropertyInfo.SetValue(item, parsedValue);
                    }
                    catch (Exception ex)
                    {
                        // Indicate parsing error on row?
                    }
                });

                retList.Add(item);
            }

            return retList;
        }
    }
}
