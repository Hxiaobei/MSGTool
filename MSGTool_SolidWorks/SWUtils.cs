
using System;
using System.Diagnostics;
using System.Linq;
using Xarial.XCad.SolidWorks;

enum SwDocumentTypes {
    swDocNONE,
    swDocPART,
    swDocASSEMBLY,
    swDocDRAWING,
    swDocSDM,
    swDocLAYOUT,
    swDocSM,
}
public class SWUtils {

    static ISwApplication ISwApp;
    public static ISwApplication GetSwApp() {

        if(ISwApp != null) return ISwApp;
        try {
            if(!IsSolidWorksRunning) return null;
            ISwApp = SwApplicationFactory.FromProcess(Process.GetProcessesByName("SLDWORKS").FirstOrDefault());
        } catch(Exception) {

            throw new Exception("failed to find SOLIDWORKS.");
        }

        return ISwApp;
    }
    public static bool IsSolidWorksRunning => Process.GetProcessesByName("SLDWORKS").Length != 0;
}


