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
            var file = new FileInfo(@"C:\repo\ingredients.xlsx");

            Parser parser = new Parser();
             List<Ingredient> loadedIngredients = await parser.LoadIngredientlData(file);


            foreach (var i in loadedIngredients)
            {
                Console.WriteLine($"{i.Id } {i.Name} {i.TotalKgCo2eq} {i.Category} calories {i.Caloriesperkg}");
            }
        }


    }
}
