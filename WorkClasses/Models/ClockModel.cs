using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BalanceApp.WorkClasses.Models
{
    public class ClockModel : IComparable<ClockModel>
    {
        public int Id { get; set; }
        public string Brand { get; set; } = "";
        public string Model { get; set; } = "";
        public int Price { set; get; }
        public int Count { get; set; }

        public int CompareTo(ClockModel? other)
        {
            if (other == null) 
                throw new ArgumentNullException("Сравниваемый объект - null");
            return Model.CompareTo(other.Model);
        }
    }
}
