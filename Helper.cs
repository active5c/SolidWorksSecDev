using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SolidWorksSecDev
{
    public static class Helper
    {

        public static bool ProcessModel(SldWorks swApp, string file, List<CustomPropertyObject> customPropertys,
            CancellationToken cancellationToken)
        {
            int warning = 0;
            int error = 0;
            try
            {
                if (cancellationToken.IsCancellationRequested) return false;
                string extension = Path.GetExtension(file);
                int type = 0;
                if (extension.ToLower().Contains("sldprt")) type = (int) swDocumentTypes_e.swDocPART;
                else type = (int)swDocumentTypes_e.swDocASSEMBLY;
                ModelDoc2 swModel = swApp.OpenDoc6(file, type, (int) swOpenDocOptions_e.swOpenDocOptions_Silent, "",
                    ref error, ref warning);
                if (error != 0) return false;
                if (swModel == null) return false;
                foreach (CustomPropertyObject item in customPropertys)
                {
                    CustomPropertyManager customPropertyManager = swModel.Extension.CustomPropertyManager[""];
                    if (item.Delete)
                    {
                        DeleteCustomPropertyValue(customPropertyManager, item.Name);
                    }
                    else
                    {
                        ReplaceCustomPropertyValue(customPropertyManager, item.Name, item.Value, item.NewValue);
                    }
                    swModel.SaveSilent();
                    swApp.QuitDoc(swModel.GetTitle());
                    swModel = null;
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }

        public static string[] GetCADFilesFromDirectory(string directory)
        {
            //check if directory exists
            if (!Directory.Exists(directory)) return null;
            //get sldprts from the directory
            string[] parts = Directory.GetFiles(directory, "*.sldprt");
            string[] assemblies = Directory.GetFiles(directory, "*.sldasm");
            List<string> files=new List<string>();
            if(parts!=null) files.AddRange(parts);
            if (assemblies != null) files.AddRange(assemblies);
            return files.ToArray();
        }

        public static bool ReplaceCustomPropertyValue(CustomPropertyManager customPropertyManager,string customProperty,string replaceable, string replacing)
        {
            //get custom property value
            string val = customPropertyManager.Get(customProperty);
            //check if the custom property exists
            if (!string.IsNullOrEmpty(val))
            {
                //Check if the replaceable sting exists in the custom property value
                if (val.ToLower().Contains(replaceable.ToLower()))
                {
                    string newVal = val.Replace(replaceable, replacing);
                    int ret=customPropertyManager.Set(customProperty, newVal);
                    if (ret == 0) return true;
                    else return false;
                }
            }
            return false;
        }

        public static bool DeleteCustomPropertyValue(CustomPropertyManager customPropertyManager, string customProperty)
        {
            //get custom property value
            string val = customPropertyManager.Get(customProperty);
            //check if the custom property exists
            if (!string.IsNullOrEmpty(val))
            {
                int ret = customPropertyManager.Delete(customProperty);
                if (ret == 0) return true;
                else return false;
            }
            return false;
        }
    }
}
