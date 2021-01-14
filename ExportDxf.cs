using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SolidWorksSecDev
{
    public class ExportDxf
    {
        ModelDoc2 swModel = default(ModelDoc2);//装配体或零件
        PartDoc swPartDoc = default(PartDoc);//零件

        AssemblyDoc swAssy = default(AssemblyDoc);//装配体
        Component2 swComp = default(Component2);//子装配或零件

        ModelDoc2 swPart = default(ModelDoc2);//子装配中的零件

        /// <summary>
        /// 遍历装配体导出钣金dxf下料图
        /// </summary>
        /// <param name="swApp"></param>
        public void AssyExportDxf(SldWorks swApp)
        {
            Dictionary<string, int> sheetMetalDic = new Dictionary<string, int>();

            swModel = (ModelDoc2)swApp.ActiveDoc;//获取当前打开的零件
            if (swModel == null)
            {
                MessageBox.Show("没有打开装配体");
                return;
            }
            //判断为装配体时继续执行，否则跳出
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY) return;
            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            //获取所有零部件集合
            var compList = swAssy.GetComponents(false);
            //遍历集合中的所有零部件对象，判断并获取需要导图的零件
            foreach (var item in compList)
            {
                swComp = (Component2)item;
                //判断需要导出下料图的零件：1.是否显示，2.是否被压缩，3.是否封套，4.是否为零件
                if (swComp.Visible != (int)swComponentVisibilityState_e.swComponentVisible
                    || swComp.IsSuppressed() || swComp.IsEnvelope()
                    || Path.GetExtension(swComp.GetPathName()).ToLower() != ".sldprt")
                    continue;//继续遍历下一个组件

                //递归判断父装配体的状态
                if (ParentCompState(swComp)) continue;

                //获取文档中的实体Body对象集合
                var bodyList = swComp.GetBodies2((int)swBodyType_e.swSolidBody);
                //遍历集合中的所有Body对象,判断是否为钣金
                foreach (var swBody in bodyList)
                {
                    //如果是钣金则将零件地址添加到字典中,存在则数量+1
                    if (!swBody.IsSheetMetal()) continue;
                    if (sheetMetalDic.ContainsKey(swComp.GetPathName())) sheetMetalDic[swComp.GetPathName()] += 1;
                    else sheetMetalDic.Add(swComp.GetPathName(), 1);
                }

            }
            string assyPath = swModel.GetPathName();
            //关闭装配体零件
            swApp.CloseDoc(assyPath);
            string dxfPath = assyPath.Substring(0, assyPath.Length - 7) + @"-DXF\";
            //判断文件夹是否存在，不存在就创建它
            if (!Directory.Exists(dxfPath)) Directory.CreateDirectory(dxfPath);

            //遍历钣金零件
            foreach (var sheetMetal in sheetMetalDic)
            {
                int errors = 0;
                int warnings = 0;
                //打开模型
                swPart = swApp.OpenDoc6(sheetMetal.Key, (int)swDocumentTypes_e.swDocPART,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
                //导图
                swPartDoc = (PartDoc)swPart;
                string swModelName = swPart.GetPathName(); ;//零件地址
                string swModelTitle = swPart.GetTitle();
                //带后缀的情况
                string swDxfName = dxfPath + swModelTitle.Substring(0, swModelTitle.Length - 7) + ".dxf";//Dxf图地址,或者dwg文件
                //判断不带后缀的情况
                if (swModelTitle.Substring(swModelTitle.Length - 7).ToLower() != ".sldprt")
                    swDxfName = dxfPath + swModelTitle + ".dxf";
                //导出零件
                PartExportDxf(swPartDoc, swDxfName, swModelName);
                //关闭零件
                swApp.CloseDoc(sheetMetal.Key);
            }
            //清除字典
            sheetMetalDic.Clear();
        }

        /// <summary>
        /// 零件导出dxf图方法
        /// </summary>
        public void PartExportDxf(PartDoc swPartDoc, string swDxfName, string swModelName)
        {
            object varAlignment;//dxf图方向（钣金拉丝方向），下一节课讲
            double[] dataAlignment = new double[12];
            int options;

            dataAlignment[0] = 0.0;
            dataAlignment[1] = 0.0;
            dataAlignment[2] = 0.0;
            dataAlignment[3] = 1.0;
            dataAlignment[4] = 0.0;
            dataAlignment[5] = 0.0;
            dataAlignment[6] = 0.0;
            dataAlignment[7] = 1.0;
            dataAlignment[8] = 0.0;
            dataAlignment[9] = 0.0;
            dataAlignment[10] = 0.0;
            dataAlignment[11] = 1.0;

            //预先绘制3D草图，重命名为xy，长边作为X轴，短边作为Y轴，用于限定拉丝方向
            bool status = false;
            ModelDoc2 swPart2 = (ModelDoc2)swPartDoc;
            if (swPart2.Extension.SelectByID2("xy", "SKETCH", 0, 0, 0, false, 0, null, 0)) status = true;
            if (status)
            {
                Feature swFeature = swPart2.SelectionManager.GetSelectedObject6(1, -1);
                Sketch swSketch = swFeature.GetSpecificFeature2();
                var swSketchPoints = swSketch.GetSketchPoints2();
                //获取草图中的所有点
                //用这三个点计算构成的线的长度，并判断长度，长边作为X轴，
                //画3D草图的时候一次性画出两条线，保证画点的顺序，否则会判断错误
                SketchPoint p0 = swSketchPoints[0];//最先画的点
                SketchPoint p1 = swSketchPoints[1];//第二点作为坐标原点
                SketchPoint p2 = swSketchPoints[2];//最后画的点
                                                   //原点p1
                dataAlignment[0] = p1.X;
                dataAlignment[1] = p1.Y;
                dataAlignment[2] = p1.X;
                //判断长短确定XY方向
                double l1 = Math.Sqrt(Math.Pow(p0.X - p1.X, 2) + Math.Pow(p0.Y - p1.Y, 2) + Math.Pow(p0.Z - p1.Z, 2));
                double l2 = Math.Sqrt(Math.Pow(p2.X - p1.X, 2) + Math.Pow(p2.Y - p1.Y, 2) + Math.Pow(p2.Z - p1.Z, 2));
                if (l1 > l2)
                {
                    //p1p0长，定义为x轴,p1p2短，定义为y轴
                    dataAlignment[3] = p0.X - p1.X;
                    dataAlignment[4] = p0.Y - p1.Y;
                    dataAlignment[5] = p0.Z - p1.Z;
                    dataAlignment[6] = p2.X - p1.X;
                    dataAlignment[7] = p2.Y - p1.Y;
                    dataAlignment[8] = p2.Z - p1.Z;
                }
                else
                {
                    //p1p2长，定义为x轴,p1p0短，定义为y轴
                    dataAlignment[3] = p2.X - p1.X;
                    dataAlignment[4] = p2.Y - p1.Y;
                    dataAlignment[5] = p2.Z - p1.Z;
                    dataAlignment[6] = p0.X - p1.X;
                    dataAlignment[7] = p0.Y - p1.Y;
                    dataAlignment[8] = p0.Z - p1.Z;
                }
            }
            varAlignment = dataAlignment;

            //Export sheet metal to a single drawing file，将钣金导出到单个工程图文件
            options = 1;  //include flat-pattern geometry
            swPartDoc.ExportToDWG2(swDxfName, swModelName, (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal, true, varAlignment, false, false, options, null);

        }

        /// <summary>
        /// 递归方法
        /// </summary>
        /// <param name="swComp"></param>
        private bool ParentCompState(Component2 swComp2)
        {
            Component2 swParentComp = swComp2.GetParent();//获取父装配体
            //直接装配在总装中的零件，GetParent方法会返回null，参见方法的remarks，此时无需判断父装配体
            //不为null，则需要判断父装配体：1.是否显示，2.是否被压缩，3.是否封套
            if (swParentComp != null)
            {
                Debug.Print(swParentComp.GetPathName());
                if (swParentComp.Visible != (int)swComponentVisibilityState_e.swComponentVisible
                    || swParentComp.IsSuppressed() || swParentComp.IsEnvelope())
                    return true;//继续遍历下一个 
                return ParentCompState(swParentComp);//递归操作
            }
            return false;
        }
    }
}
