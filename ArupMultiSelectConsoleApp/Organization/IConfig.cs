using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Organization
{
    interface IConfig
    {
        string KeyVaultPath { get; }

        string CRMHubPWKey { get; }
    }
}
