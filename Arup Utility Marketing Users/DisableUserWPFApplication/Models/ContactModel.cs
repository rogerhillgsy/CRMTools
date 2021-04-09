using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DisableUserWPFApplication.Models
{
    public class ContactModel
    {
        public Guid ID { get; set; }
        public string Name { get; set; }
        public string status { get; set; }

        public string email { get; set; }
    }
}
