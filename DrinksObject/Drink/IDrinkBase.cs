using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DrinksCaculator.DrinksObject.Drink
{
    public interface IDrinkBase
    {
        string DrinkName { get; set; }
        string Sugar { get; set; }
        string Ice { get; set; }
        int Price { get; set; }
    }
}
