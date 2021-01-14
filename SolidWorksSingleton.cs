using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidWorksSecDev
{
    public class SolidWorksSingleton
    {
        private static SldWorks swApp;
        public static SldWorks GetApplication()
        {
            if (swApp == null)
            {
                swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
                swApp.Visible = true;
                return swApp;
            }
            return swApp;
        }

        public static async Task<SldWorks> GetApplicationAsync()
        {
            if (swApp == null)
            {
                return await Task<SldWorks>.Run(() =>
                {
                    swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
                    swApp.Visible = true;
                    return swApp;
                });
            }
            return swApp;
        }

        public static void Dipose()
        {
            if (swApp != null)
            {
                swApp.ExitApp();
                swApp = null;
            }
        }
    }
}
