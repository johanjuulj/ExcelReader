using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

namespace ExcelReader
{
    class Ingredient
    {

        public string Id { get; set; } 
        public string Name { get; set; }
        //public int ShelfLifeDays { get; set; }

        public CategoryDSK Category { get; set; }
        [Column(TypeName = "decimal(18,4)")]
        public decimal TotalKgCo2eq { get; set; }


        //data annotation http://jameschambers.com/2019/06/No-Type-Was-Specified-for-the-Decimal-Column/
        [Column(TypeName = "decimal(18,4)")]
        public decimal Caloriesperkg { get; set; }
       
        public IList<RecipeIngredient> RecipeIngredients { get; set; } //non instantiated?

        public Ingredient()
        {

        }
        public Ingredient(string Id, string Name)
        {
            this.Id = Id;
            this.Name = Name;
        }
        public Ingredient(string Id, string Name, decimal CO2Per100G, decimal CaloriesPer100G)
        {
            this.Id = Id;
            this.Name = Name;
            
            this.TotalKgCo2eq = CO2Per100G;
            this.Caloriesperkg = CaloriesPer100G;
        }
    }
}
