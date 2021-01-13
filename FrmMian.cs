using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidWorks.Interop.swconst;


namespace SolidWorksSecDev
{
    public partial class FrmMian : Form
    {
        private SldWorks swApp;//SolidWorks程序

        ModelDoc2 swModel = default(ModelDoc2);//装配体或零件
        PartDoc swPartAlone = default(PartDoc);//零件
        Feature swFeat = default(Feature);//特征

        AssemblyDoc swAssy = default(AssemblyDoc);//装配体
        Component2 swComp = default(Component2);//子装配或零件

        ModelDoc2 swSubModel = default(ModelDoc2);//子装配
        AssemblyDoc swSubAssy = default(AssemblyDoc);//子装配
        ModelDoc2 swPart = default(ModelDoc2);//子装配中的零件

        public FrmMian()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 打开/链接SW
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void btnConnect_Click(object sender, EventArgs e)
        {
            //swApp = SolidWorksSingleton.GetApplication();
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            //if (swApp != null)
            //{
            //    //显示版本
            //    string msg = "This message from C#. solidworks version is " + swApp.RevisionNumber();
            //    swApp.SendMsgToUser(msg);
            //}
            StudyCode study=new StudyCode();
            study.CADSharpP4(swApp);
        }
        /// <summary>
        /// 编辑零件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void btnEditSwModel_Click(object sender, EventArgs e)
        {
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            if (swApp == null) return;
            swApp.CommandInProgress = true; //告诉SolidWorks，现在是用外部程序调用命令
            swModel = (ModelDoc2)swApp.ActiveDoc;//获取当前打开的零件
            //判断为零件时继续执行，否则跳出
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART) return;
            swPartAlone = (PartDoc)swModel;
            object configNames = null;

            try
            {

                //修改参数，注意SW中单位为米，换算成mm应当除以1000
                swModel.Parameter("D2@Sketch2").SystemValue = 200 / 1000m;
                //压缩特征
                swFeat = swPartAlone.FeatureByName("Edge-Flange1");
                swFeat.SetSuppression2(1, 2, configNames); //参数1：1解压，0压缩，2解压缩这个特征的子特征

                swModel.ForceRebuild3(true);//设置成true，直接更新顶层，速度很快，设置成false，每个零件都会更新，很慢
                swModel.Save();//保存，很耗时间
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                swApp.CommandInProgress = false; //及时关闭外部命令调用，否则影响SolidWorks的使用
            }

        }
        /// <summary>
        /// 编辑装配体
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void btnEditSwAssy_Click(object sender, EventArgs e)
        {
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            if (swApp == null) return;
            swApp.CommandInProgress = true; //告诉SolidWorks，现在是用外部程序调用命令
            swModel = (ModelDoc2)swApp.ActiveDoc;//获取当前打开的零件/装配体
            //判断不是装配体直接跳出
            //if (Path.GetExtension(swModel.GetPathName()).ToUpper()!=".SLDASM") return;
            if(swModel.GetType()!=(int)swDocumentTypes_e.swDocASSEMBLY)return;
            swAssy = (AssemblyDoc)swModel;
            object configNames = null;

            try
            {
                //修改装配体顶层参数
                swModel.Parameter("D1@Distance1").SystemValue = 300 / 1000m;
                //修改装配体顶层特征
                swFeat = swAssy.FeatureByName("LocalLPattern1");
                swFeat.SetSuppression2(0, 2, configNames); //参数1：1解压，0压缩

                //压缩装配体中的零件
                swComp = swAssy.GetComponentByName("Part1-3");
                swComp.SetSuppression2(0); //2解压缩(包含子装配内部)，0压缩，3只解压子装配本身. .

                swModel.ForceRebuild3(true);//设置成true，直接更新顶层，速度很快，设置成false，每个零件都会更新，很慢
                swModel.Save();//保存，很耗时间
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                swApp.CommandInProgress = false; //及时关闭外部命令调用，否则影响SolidWorks的使用
            }
        }
        /// <summary>
        /// 编辑子装配
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void btnEditSubAssy_Click(object sender, EventArgs e)
        {
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            if (swApp == null) return;
            swApp.CommandInProgress = true; //告诉SolidWorks，现在是用外部程序调用命令
            swModel = (ModelDoc2)swApp.ActiveDoc;//获取当前打开的零件/装配体
            //判断不是装配体直接跳出
            //if (Path.GetExtension(swModel.GetPathName()).ToUpper() != ".SLDASM") return;
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY) return;
            swAssy = (AssemblyDoc)swModel;
            object configNames = null;

            try
            {
                swComp = swAssy.GetComponentByName("Assem1-1");//获取子装配
                swSubModel = swComp.GetModelDoc2(); //打开零件
                swSubModel.Parameter("D1@LocalLPattern1").SystemValue = 3;//阵列数量
                swSubModel.Parameter("D3@LocalLPattern1").SystemValue = 500 / 1000m;//阵列距离

                swSubAssy = (AssemblyDoc)swSubModel;
                swComp = swSubAssy.GetComponentByName("Part2-1");
                swPart = swComp.GetModelDoc2(); //打开零件
                swPart.Parameter("D7@边线-法兰2").SystemValue = 100 / 1000m;
                swFeat = swComp.FeatureByName("Cut-Extrude1");
                swFeat.SetSuppression2(0, 2, configNames); //参数1：1解压，0压缩

                swModel.ForceRebuild3(true);//设置成true，直接更新顶层，速度很快，设置成false，每个零件都会更新，很慢
                swModel.Save();//保存，很耗时间
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                swApp.CommandInProgress = false; //及时关闭外部命令调用，否则影响SolidWorks的使用
            }
        }
        /// <summary>
        /// 导出钣金零件下料图/设定拉丝方向
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void btnExportDxf_Click(object sender, EventArgs e)
        {
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            if (swApp == null) return;
            swApp.CommandInProgress = true; //告诉SolidWorks，现在是用外部程序调用命令
            swModel = (ModelDoc2)swApp.ActiveDoc;//获取当前打开的零件
            //判断为零件时继续执行，否则跳出
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART) return;
            swPartAlone = (PartDoc)swModel;

            string swModelName = swModel.GetPathName(); ;//零件地址
            string swDxfName= swModelName.Substring(0, swModelName.Length - 6)+"dxf";//Dxf图地址,或者dwg文件
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
            try
            {

                //预先绘制3D草图，重命名为xy，长边作为X轴，短边作为Y轴，用于限定拉丝方向
                bool status = false;
                if (swModel.Extension.SelectByID2("xy", "SKETCH", 0, 0, 0, false, 0, null, 0)) status = true;
                if (status)
                {
                    Feature swFeature = swModel.SelectionManager.GetSelectedObject6(1, -1);
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
                swPartAlone.ExportToDWG2(swDxfName, swModelName, (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal, true, varAlignment, false, false, options, null);

                MessageBox.Show("导出完成！");

                //swModel.ForceRebuild3(true);//设置成true，直接更新顶层，速度很快，设置成false，每个零件都会更新，很慢
                //swModel.Save();//保存，很耗时间
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                swApp.CommandInProgress = false; //及时关闭外部命令调用，否则影响SolidWorks的使用
            }
        }
    }
}
