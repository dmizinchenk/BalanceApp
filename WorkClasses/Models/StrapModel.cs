using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BalanceApp.WorkClasses.Models
{
    public class StrapModel : IComparable<StrapModel>
    {
        public int Id { get; set; }
        public int? Number { get; set; } = null;
        public string Name { get; set; } = "";
        public string? Name1С { get; set; }
        public int Price { set; get; }
        public int Count { get; set; }

        public int CompareTo(StrapModel? other)
        {
            if (other == null) 
                throw new ArgumentNullException("Сравниваемый объект - null");
            return Id.CompareTo(other.Id);
        }
    }
}
