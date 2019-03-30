using System;
using System.Collections.Generic;
using System.Linq;
using gMotoImpEx.Core.DomainServices;

namespace ExcImp
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var tmp = new ExcelDataSource();
                var rez = tmp.GetData<DataItem>(@"D:\REPO\_Learn\ExcImp\InputData.xlsx", new Dictionary<string, string>()
                {
                    {"įmonės kodas", "Code"},
                    {"TP valst. Nr", "Number"},
                    {"TP VIN kodas", "Vin"}
                });
                
                var dtList = rez.Where(s => !string.IsNullOrWhiteSpace(s.Code)
                                            && !string.IsNullOrWhiteSpace(s.Number)
                                            && !string.IsNullOrWhiteSpace(s.Vin))
                    .Distinct()
                    .ToList();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
