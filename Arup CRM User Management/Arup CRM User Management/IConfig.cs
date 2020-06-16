using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Arup_CRM_User_Management
{
    interface IConfig
    {
        string KeyVaultPath { get; }

        string CRMHubPWKey { get; }
    }
}
