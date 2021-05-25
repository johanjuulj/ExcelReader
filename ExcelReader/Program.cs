using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace ExcelReader
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(@"C:\repo\ingredients.xlsx");

            
            //var ingredients = GetSetupdata();

            //await SaveExcelFile(ingredients, file);

            //List<Ingredient> loadedIngredients = await LoadExcelData(file);

            List<Ingredient> loadedIngredients = await LoadIngredientlData(file);

            

            foreach (var i in loadedIngredients)
            {
                Console.WriteLine($"{i.id } {i.name}");
            }
        }

        private static async Task<List<Ingredient>> LoadIngredientlData(FileInfo file)
        {
            List<Ingredient> output = new List<Ingredient>();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var ws = package.Workbook.Worksheets[1];

            int row = 2;
            int col = 1;
            
            //while statement checks whether the field is null. if it is the document is over
            while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
            {
                Console.WriteLine("Not not null");
                Ingredient p = new Ingredient

                {

                    
                    id = ws.Cells[row, col].Value.ToString(),
                    name = ws.Cells[row, col + 1].Value.ToString(),
                    //co2 = int.Parse(ws.Cells[row, col + 12].Value.ToString())



                };
                output.Add(p);
                row += 1;
                Console.WriteLine("Not not null");

            }
            return output;
        }

        private static async Task<List<Ingredient>> LoadExcelData(FileInfo file)
        {
            List<Ingredient> output = new List<Ingredient>();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var ws = package.Workbook.Worksheets[0];

            int row = 2;
            int col = 1;

            //while statement checks whether the field is null. if it is the document is over
            while (string.IsNullOrWhiteSpace(ws.Cells[row,col].Value?.ToString()) == false)
            {
                Ingredient p = new Ingredient

                {
                    id = ws.Cells[row, col].Value.ToString(),
                    //name = ws.Cells[row, col + 1].Value.ToString(),
                    //co2 = int.Parse(ws.Cells[row, col + 2].Value.ToString())
                    


                };
                output.Add(p);
                row += 1;

            
            }
            return output;
        }

        private static async Task SaveExcelFile(List<Ingredient> ingredients, FileInfo file)
        {
            //remove after testing
            DeleteIfExists(file);

            using (var package = new ExcelPackage(file))
            {

                var ws = package.Workbook.Worksheets.Add("MainReport");

                var range = ws.Cells["A1"].LoadFromCollection(ingredients, true);
                range.AutoFitColumns();

                //styling
                ws.Cells["A1"].Value = "Carbon Overview of Ingredients";
                ws.Cells["A1:C1"].Merge = true;
                ws.Column(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Row(1).Style.Font.Size = 18;
                ws.Row(1).Style.Font.Color.SetColor(Color.Blue);

                ws.Row(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Column(3).Width = 20;

                await package.SaveAsync();
            }
            
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        //private static List<Ingredient> GetSetupdata()
        //{
        //    List<Ingredient> output = new List<Ingredient>()
        //    {
        //        new Ingredient() { id = "1", name = "Gulerod", co2= 2},
        //        new Ingredient() { id = "2", name = "Orangerod", co2= 7},
        //        new Ingredient() { id = "3", name = "Hakket Oksekød", co2= 200},
        //        new Ingredient() { id = "4", name = "Tun i Vand", co2= 65}
        //    };
        //    return output;
        //}

        public class Ingredient
        {
            public string id { get; set; }
            public string name { get; set; }
            //public int co2 { get; set; }

        }
    }
}
