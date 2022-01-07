using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Crestron.SimplSharpPro;

namespace FTR_Utils
{ 
    public class FTR_Utils
    {
        PanelSigInfo psi;
        Sheet sheet;
        eSigType sig_type = new eSigType(); 

        public FTR_Utils()
        {

            psi = new PanelSigInfo(sig_type, 1, 0x20);
            sheet = new Sheet();
        }
    }
}
