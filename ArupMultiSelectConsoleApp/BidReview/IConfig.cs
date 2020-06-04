using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BidReview
{
    public interface IConfig
    {
        string KeyVaultPath { get; }

        string CRMHubPWKey { get; }
    }
}
