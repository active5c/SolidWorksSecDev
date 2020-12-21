using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace SolidWorksSecDev
{
    public partial class FrmMian : Form
    {
        private SldWorks swApp;
        ModelDoc2 swModel = default(ModelDoc2);
        PartDoc swPart = default(PartDoc);
        Feature swFeat = default(Feature);
        AssemblyDoc swAssy = default(AssemblyDoc);//装配体
        Component2 swComp = default(Component2);//子装配或零件


        public FrmMian()
        {
            InitializeComponent();
        }

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
        }

        private async void btnEditSwModel_Click(object sender, EventArgs e)
        {
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            if (swApp == null) return;
            swApp.CommandInProgress = true; //告诉SolidWorks，现在是用外部程序调用命令
            swModel = (ModelDoc2)swApp.ActiveDoc;//获取当前打开的零件
            swPart = (PartDoc)swModel;
            object configNames = null;

            try
            {

                //修改参数，注意SW中单位为米，换算成mm应当除以1000
                swModel.Parameter("D2@Sketch2").SystemValue = 200 / 1000m;
                //压缩特征
                swFeat = swPart.FeatureByName("Edge-Flange1");
                swFeat.SetSuppression2(2, 2, configNames); //参数1：1解压，0压缩

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

        private async void btnEditSwAssy_Click(object sender, EventArgs e)
        {
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            if (swApp == null) return;
            swApp.CommandInProgress = true; //告诉SolidWorks，现在是用外部程序调用命令
            swModel = (ModelDoc2)swApp.ActiveDoc;//获取当前打开的零件/装配体
            //if(swModel.)
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
    }
}
