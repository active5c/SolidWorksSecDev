namespace SolidWorksSecDev
{
    partial class FrmMian
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnConnect = new System.Windows.Forms.Button();
            this.btnEditSwModel = new System.Windows.Forms.Button();
            this.btnEditSwAssy = new System.Windows.Forms.Button();
            this.btnEditSubAssy = new System.Windows.Forms.Button();
            this.btnExportDxf = new System.Windows.Forms.Button();
            this.btnTraverseAssy = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnConnect
            // 
            this.btnConnect.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnConnect.Location = new System.Drawing.Point(12, 12);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(345, 45);
            this.btnConnect.TabIndex = 0;
            this.btnConnect.Text = "1.打开/连接SW+学习代码测试";
            this.btnConnect.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnConnect.UseVisualStyleBackColor = true;
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // btnEditSwModel
            // 
            this.btnEditSwModel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEditSwModel.Location = new System.Drawing.Point(12, 64);
            this.btnEditSwModel.Name = "btnEditSwModel";
            this.btnEditSwModel.Size = new System.Drawing.Size(111, 42);
            this.btnEditSwModel.TabIndex = 0;
            this.btnEditSwModel.Text = "2.修改零件";
            this.btnEditSwModel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEditSwModel.UseVisualStyleBackColor = true;
            this.btnEditSwModel.Click += new System.EventHandler(this.btnEditSwModel_Click);
            // 
            // btnEditSwAssy
            // 
            this.btnEditSwAssy.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEditSwAssy.Location = new System.Drawing.Point(129, 64);
            this.btnEditSwAssy.Name = "btnEditSwAssy";
            this.btnEditSwAssy.Size = new System.Drawing.Size(111, 42);
            this.btnEditSwAssy.TabIndex = 1;
            this.btnEditSwAssy.Text = "3.修改装配体";
            this.btnEditSwAssy.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEditSwAssy.UseVisualStyleBackColor = true;
            this.btnEditSwAssy.Click += new System.EventHandler(this.btnEditSwAssy_Click);
            // 
            // btnEditSubAssy
            // 
            this.btnEditSubAssy.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEditSubAssy.Location = new System.Drawing.Point(246, 64);
            this.btnEditSubAssy.Name = "btnEditSubAssy";
            this.btnEditSubAssy.Size = new System.Drawing.Size(111, 42);
            this.btnEditSubAssy.TabIndex = 1;
            this.btnEditSubAssy.Text = "4.修改子装配体";
            this.btnEditSubAssy.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEditSubAssy.UseVisualStyleBackColor = true;
            this.btnEditSubAssy.Click += new System.EventHandler(this.btnEditSubAssy_Click);
            // 
            // btnExportDxf
            // 
            this.btnExportDxf.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExportDxf.Location = new System.Drawing.Point(12, 118);
            this.btnExportDxf.Name = "btnExportDxf";
            this.btnExportDxf.Size = new System.Drawing.Size(111, 42);
            this.btnExportDxf.TabIndex = 1;
            this.btnExportDxf.Text = "5-6.导出钣金dxf图";
            this.btnExportDxf.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExportDxf.UseVisualStyleBackColor = true;
            this.btnExportDxf.Click += new System.EventHandler(this.btnExportDxf_Click);
            // 
            // btnTraverseAssy
            // 
            this.btnTraverseAssy.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnTraverseAssy.Location = new System.Drawing.Point(129, 118);
            this.btnTraverseAssy.Name = "btnTraverseAssy";
            this.btnTraverseAssy.Size = new System.Drawing.Size(111, 42);
            this.btnTraverseAssy.TabIndex = 1;
            this.btnTraverseAssy.Text = "7.遍历装配体导图";
            this.btnTraverseAssy.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnTraverseAssy.UseVisualStyleBackColor = true;
            this.btnTraverseAssy.Click += new System.EventHandler(this.btnTraverseAssy_Click);
            // 
            // FrmMian
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(368, 505);
            this.Controls.Add(this.btnTraverseAssy);
            this.Controls.Add(this.btnExportDxf);
            this.Controls.Add(this.btnEditSubAssy);
            this.Controls.Add(this.btnEditSwAssy);
            this.Controls.Add(this.btnEditSwModel);
            this.Controls.Add(this.btnConnect);
            this.Name = "FrmMian";
            this.Text = "SoliWorks二次开发专题讲解";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnConnect;
        private System.Windows.Forms.Button btnEditSwModel;
        private System.Windows.Forms.Button btnEditSwAssy;
        private System.Windows.Forms.Button btnEditSubAssy;
        private System.Windows.Forms.Button btnExportDxf;
        private System.Windows.Forms.Button btnTraverseAssy;
    }
}

