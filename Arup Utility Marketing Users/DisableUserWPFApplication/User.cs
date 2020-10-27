using System;
using System.Collections.Generic;

namespace DisableUserWPFApplication
{
    public class User
    {
        public string DomainUserName { get; set; }
        public Guid UserId { get; set; }
        public int LicenseType { get; set; }
        public string FullName { get; set; }
        public bool status { get; set; }
        public List<Guid> roles { get; set; }
    }
}