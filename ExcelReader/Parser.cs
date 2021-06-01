using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    public class Parser
    {

        public Parser()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            
            

            //var ingredients = GetSetupdata();

            //await SaveExcelFile(ingredients, file);

            //List<Ingredient> loadedIngredients = await LoadExcelData(file);
        }
        public  async Task<List<Ingredient>> LoadIngredientlData(FileInfo file)
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

                Ingredient newI = new Ingredient

                {


                    Id = ws.Cells[row, col].Value.ToString(),
                    Name = ws.Cells[row, col + 3].Value.ToString(),
                    TotalKgCo2eq = decimal.Parse(ws.Cells[row, col + 12].Value.ToString()),
                    Caloriesperkg = decimal.Parse(ws.Cells[row, col + 13].Value.ToString()),




                };

                string category = ws.Cells[row, col + 4].Value.ToString();


                if (category == CategoryDSK.Vegetables.ToString())
                {
                    newI.Category = CategoryDSK.Vegetables;
                }
                if (category == "Meat/poultry")
                {
                    newI.Category = CategoryDSK.Meat_poultry;
                }
                if (category == CategoryDSK.Beverages.ToString())
                {
                    newI.Category = CategoryDSK.Beverages;
                }
                if (category == "Bread/bakery products")
                {
                    newI.Category = CategoryDSK.Bread_bakeryProducts;
                }
                if (category == "Cereal/grain/pulse products")
                {
                    newI.Category = CategoryDSK.Cereal_grain_pulseProducts;
                }
                if (category == "Milk/eggs/substitute products")
                {
                    newI.Category = CategoryDSK.Milk_eggs_substituteProducts;
                }
                if (category == "Seasonings/preservatives/extracts")
                {
                    newI.Category = CategoryDSK.Seasonings_preservatives_extracts;
                }
                if (category == CategoryDSK.Seafood.ToString())
                {
                    newI.Category = CategoryDSK.Seafood;
                }
                if (category == "Prepared/preserved foods")
                {
                    newI.Category = CategoryDSK.Prepared_preservedFoods;
                }
                if (category == "Candy/sugar products")
                {
                    newI.Category = CategoryDSK.Candy_sugerProducts;
                }

                if (category == CategoryDSK.Fruits.ToString())
                {
                    newI.Category = CategoryDSK.Fruits;
                }
                if (category == "Fruit/vegetable products")
                {
                    newI.Category = CategoryDSK.Fruit_vegetableProducts;
                }
                if (category == "Oils/fats edible")
                {
                    newI.Category = CategoryDSK.Oils_fatsEdible;
                }
                else
                {

                }

                output.Add(newI);
                row += 1;


            }
            return output;
        }

        public  async Task<List<Ingredient>> LoadExcelData(FileInfo file)
        {
            List<Ingredient> output = new List<Ingredient>();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var ws = package.Workbook.Worksheets[0];

            int row = 2;
            int col = 1;

            //while statement checks whether the field is null. if it is the document is over
            while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
            {
                Ingredient p = new Ingredient

                {
                    Id = ws.Cells[row, col].Value.ToString(),
                    Name = ws.Cells[row, col + 1].Value.ToString(),
                    //co2 = int.Parse(ws.Cells[row, col + 2].Value.ToString())



                };
                output.Add(p);
                row += 1;


            }
            return output;
        }

        public  async Task SaveExcelFile(List<Ingredient> ingredients, FileInfo file)
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

        public  void DeleteIfExists(FileInfo file)
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

        //public class Ingredient
        //{
        //    public string id { get; set; }
        //    public string name { get; set; }
        //    //public int co2 { get; set; }

        //}
    }
}

