using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Xarial.XCad.SolidWorks;

static class SldWorksEx {

    public static ISldWorks GetSwApp() => GetISwApp()?.Sw;

    static ISwApplication ISwApp;
    public static ISwApplication GetISwApp() {

        if(ISwApp == null)
            ISwApp = SwApplicationFactory.FromProcess(Process.GetProcessesByName("SLDWORKS").FirstOrDefault());

        return ISwApp;
    }
}

