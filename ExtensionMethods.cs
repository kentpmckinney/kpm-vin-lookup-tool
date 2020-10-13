using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VehicleInformationLookupTool
{
    public static class Extension
    {
        public static void ThrowIfNullOrEmpty(this object target)
        {
            switch (target)
            {
                case null:
                {
                    throw new ArgumentNullException(nameof(target));
                }

                case string s:
                {
                    if (string.IsNullOrWhiteSpace(s))
                    {
                        throw new ArgumentNullException(nameof(target));
                    }
                    break;
                }
            }
        }
    }
}
