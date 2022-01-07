using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Crestron.SimplSharpPro;

namespace FTR_Utils
{
    public class PanelSigInfo
    {
        public PanelSigInfo(eSigType _type, uint _n, uint _smartID)
        {
            eSigType type = _type;
            uint n = _n;
            uint smartID = _smartID;
        }
    }
}
