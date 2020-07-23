using DrinksCaculator.DrinksObject.Member;

namespace DrinksCaculator.DrinksObject.Drink
{
    public class Drink : IMemberBase, IDrinkBase
    {
        public string No { get; set; }

        public string StaffName { get; set; }

        public string DrinkName { get; set; }

        public string Sugar { get; set; }

        public string Ice { get; set; }

        public int Price { get; set; }

        public string Remark { get; set; }
    }
}
