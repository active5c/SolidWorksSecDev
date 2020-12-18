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
        public FrmMian()
        {
            InitializeComponent();
        }

        private async void btnConnect_Click(object sender, EventArgs e)
        {
            //swApp = SolidWorksSingleton.GetApplication();
            swApp = await SolidWorksSingleton.GetApplicationAsync();
            if (swApp != null)
            {
                //显示版本
                string msg = "This message from C#. solidworks version is " + swApp.RevisionNumber();
                swApp.SendMsgToUser(msg);
            }
        }


    }
}
