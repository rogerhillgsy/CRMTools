using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DisableUserWPFApplication.Models
{
    public class MarketingListModel
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string GuidValue { get; set; }

        public override string ToString()
        {
            return this.Name.ToString();
        }
    }
}
