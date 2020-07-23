using DrinksCaculator.DrinksObject.Drink;

namespace DrinksCaculator.DrinksObject.Order
{
    public class Order : IDrinkBase
    {
        public string DrinkName { get; set; }

        public string Sugar { get; set; }

        public string Ice { get; set; }

        public int Price { get; set; }

        public int SumOfCups { get; set; }

        public string[] PoepleOfOrdered { get; set; }
    }
}
