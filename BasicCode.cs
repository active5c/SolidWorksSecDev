using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SolidWorksSecDev
{
    public class BasicCode
    {
        /// <summary>
        /// 编辑零件
        /// </summary>
        /// <param name="swApp">SW程序</param>
        public void EditSwModel(SldWorks swApp)
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;//获取当前打开的零件
            //判断为零件时继续执行，否则跳出
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART) return;
            PartDoc swPartDoc = (PartDoc)swModel;
            object configNames = null;

            //修改参数，注意SW中单位为米，换算成mm应当除以1000
            swModel.Parameter("D2@Sketch2").SystemValue = 200 / 1000m;
            //压缩特征
            Feature swFeat = swPartDoc.FeatureByName("Edge-Flange1");
            swFeat.SetSuppression2(1, 2, configNames); //参数1：1解压，0压缩，2解压缩这个特征的子特征

            //设置成true，直接更新顶层，速度很快，设置成false，每个零件都会更新，很慢
            swModel.ForceRebuild3(true);
            swModel.Save();//保存，很耗时间
        }
        /// <summary>
        ///  编辑装配体
        /// </summary>
        /// <param name="swApp">SW程序</param>
        public void EditSwAssy(SldWorks swApp)
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;//获取当前打开的零件/装配体
            //判断不是装配体直接跳出
            //if (Path.GetExtension(swModel.GetPathName()).ToUpper()!=".SLDASM") return;
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY) return;
            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            object configNames = null;

            //修改装配体顶层参数
            swModel.Parameter("D1@Distance1").SystemValue = 300 / 1000m;
            //修改装配体顶层特征
            Feature swFeat = swAssy.FeatureByName("LocalLPattern1");
            swFeat.SetSuppression2(0, 2, configNames); //参数1：1解压，0压缩
            //压缩装配体中的零件
            Component2 swComp = swAssy.GetComponentByName("Part1-3");
            swComp.SetSuppression2(0); //2解压缩(包含子装配内部)，0压缩，3只解压子装配本身. .

            swModel.ForceRebuild3(true);
            swModel.Save();//保存，很耗时间
        }
        /// <summary>
        ///  编辑子装配体
        /// </summary>
        /// <param name="swApp">SW程序</param>
        public void EditSubAssy(SldWorks swApp)
        {
            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;//获取当前打开的零件/装配体
            //判断不是装配体直接跳出
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY) return;
            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            object configNames = null;

            Component2 swComp = swAssy.GetComponentByName("Assem1-1");//获取子装配
            ModelDoc2 swSubModel = swComp.GetModelDoc2(); //打开零件
            swSubModel.Parameter("D1@LocalLPattern1").SystemValue = 3;//阵列数量
            swSubModel.Parameter("D3@LocalLPattern1").SystemValue = 500 / 1000m;//阵列距离

            AssemblyDoc swSubAssy = (AssemblyDoc)swSubModel;
            swComp = swSubAssy.GetComponentByName("Part2-1");
            ModelDoc2 swPart = swComp.GetModelDoc2(); //打开零件
            swPart.Parameter("D7@边线-法兰2").SystemValue = 100 / 1000m;
            Feature swFeat = swComp.FeatureByName("Cut-Extrude1");
            swFeat.SetSuppression2(0, 2, configNames); //参数1：1解压，0压缩

            swModel.ForceRebuild3(true);
            swModel.Save();//保存，很耗时间
        }
    }
}
