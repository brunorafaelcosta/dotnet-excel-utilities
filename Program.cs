using System;
using System.Collections.Generic;
using System.Linq;

namespace dotnet_excel_utilities
{
    class Program
    {
        const string filePath = "RegionsExport.xlsx";
        static void Main(string[] args)
        {
            bool isToImport = true;

            if (args != null && args.Length > 0)
            {
                if (args[0] == "i")
                    isToImport = true;
                else
                    isToImport = false;
            }
            
            var exportData = DataSet.GetData();

            if (isToImport)
            {
                var importedData = new ExcelUtilities().Import<DataSet.Region>(filePath);
                
                int numberOfErrors = GetNumberOfErrorsFromImportedData(exportData, importedData);

                Console.WriteLine(numberOfErrors == 0 ? "Excel file successfully imported!" : $"Excel file imported with {numberOfErrors} errors!");
            }
            else
            {
                ExcelUtilities.Export(
                    exportData
                    , filePath);

                Console.WriteLine("Excel file successfully exported!");
            }

            Console.ReadKey();
        }
    
        private static int GetNumberOfErrorsFromImportedData(IEnumerable<DataSet.Region> exportData, IEnumerable<DataSet.Region> importedData)
        {
            int numberOfErrors = 0;

            foreach (var region in exportData)
            {
                if (importedData.Any(r => r.Name == region.Name))
                {
                    var importedRegion = importedData.Single(r => r.Name == region.Name);
                    foreach (var country in region.Countries)
                    {
                        if (importedRegion.Countries.Any(c => c.Name == country.Name))
                        {
                            var importedCountry = importedRegion.Countries.Single(c => c.Name == country.Name);

                            if (importedCountry.Area != country.Area)
                                numberOfErrors++;
                            if (importedCountry.Population != country.Population)
                                numberOfErrors++;
                        }
                        else
                        {
                            numberOfErrors++;
                        }
                    }
                }
                else
                {
                    numberOfErrors++;
                }
            }

            return numberOfErrors;
        }
    }
}
