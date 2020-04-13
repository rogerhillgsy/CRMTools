using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LeadNew
{
    interface IConfig
    {
        string KeyVaultPath { get; }

        string CRMHubPWKey { get; }
    }
}
