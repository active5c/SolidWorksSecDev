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
        public FrmMian()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 打开/链接SW
        /// </summary>
        private async void btnConnect_Click(object sender, EventArgs e)
        {
            //普通方式
            //swApp = SolidWorksSingleton.GetApplication();
            //异步方式
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            if (swApp == null) return;
            //if (swApp != null) swApp.SendMsgToUser("Solidworks version: "+ swApp.RevisionNumber());

            //学习代码测试
            StudyCode study = new StudyCode();
            study.SolidWorksAcademy2019P1(swApp);
        }

        /// <summary>
        /// 编辑零件
        /// </summary>
        private async void btnEditSwModel_Click(object sender, EventArgs e)
        {
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            if (swApp == null) return;
            swApp.CommandInProgress = true; //告诉SolidWorks，现在是用外部程序调用命令

            BasicCode basicCode = new BasicCode();
            basicCode.EditSwModel(swApp);

            swApp.CommandInProgress = false; //及时关闭外部命令调用，否则影响SolidWorks的使用
            MessageBox.Show("修改零件完成");
        }

        /// <summary>
        /// 编辑装配体
        /// </summary>
        private async void btnEditSwAssy_Click(object sender, EventArgs e)
        {
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            if (swApp == null) return;
            swApp.CommandInProgress = true; //告诉SolidWorks，现在是用外部程序调用命令

            BasicCode basicCode = new BasicCode();
            basicCode.EditSwAssy(swApp);

            swApp.CommandInProgress = false; //及时关闭外部命令调用，否则影响SolidWorks的使用
            MessageBox.Show("修改装配体完成");
        }

        /// <summary>
        /// 编辑子装配体
        /// </summary>
        private async void btnEditSubAssy_Click(object sender, EventArgs e)
        {
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            if (swApp == null) return;
            swApp.CommandInProgress = true; //告诉SolidWorks，现在是用外部程序调用命令

            BasicCode basicCode = new BasicCode();
            basicCode.EditSubAssy(swApp);

            swApp.CommandInProgress = false; //及时关闭外部命令调用，否则影响SolidWorks的使用
            MessageBox.Show("修改子装配体完成");
        }

        /// <summary>
        /// 导出钣金零件下料图/设定拉丝方向
        /// </summary>
        private async void btnExportDxf_Click(object sender, EventArgs e)
        {
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            if (swApp == null) return;
            swApp.CommandInProgress = true; //告诉SolidWorks，现在是用外部程序调用命令
            
            ExportDxf exportDxf = new ExportDxf();
            exportDxf.PartExportDxf(swApp);

            swApp.CommandInProgress = false; //及时关闭外部命令调用，否则影响SolidWorks的使用
            MessageBox.Show("导出钣金下料图完成！");
        }

        /// <summary>
        /// 遍历装配体导出钣金下料图
        /// </summary>
        private async void btnTraverseAssy_Click(object sender, EventArgs e)
        {
            //链接SolidWorks
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            if (swApp == null) return;
            swApp.CommandInProgress = true; //告诉SolidWorks，现在是用外部程序调用命令

            //实例化导图类，调用导图方法
            ExportDxf exportDxf = new ExportDxf();
            exportDxf.AssyExportDxf(swApp);

            swApp.CommandInProgress = false;
            MessageBox.Show("遍历装配体导出钣金下料图完成");
        }



    }
}
