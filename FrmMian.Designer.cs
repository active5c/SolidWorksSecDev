﻿namespace SolidWorksSecDev
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
            this.SuspendLayout();
            // 
            // btnConnect
            // 
            this.btnConnect.Location = new System.Drawing.Point(32, 36);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(122, 45);
            this.btnConnect.TabIndex = 0;
            this.btnConnect.Text = "打开/连接SW";
            this.btnConnect.UseVisualStyleBackColor = true;
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // btnEditSwModel
            // 
            this.btnEditSwModel.Location = new System.Drawing.Point(170, 36);
            this.btnEditSwModel.Name = "btnEditSwModel";
            this.btnEditSwModel.Size = new System.Drawing.Size(122, 45);
            this.btnEditSwModel.TabIndex = 0;
            this.btnEditSwModel.Text = "修改零件";
            this.btnEditSwModel.UseVisualStyleBackColor = true;
            this.btnEditSwModel.Click += new System.EventHandler(this.btnEditSwModel_Click);
            // 
            // btnEditSwAssy
            // 
            this.btnEditSwAssy.Location = new System.Drawing.Point(310, 38);
            this.btnEditSwAssy.Name = "btnEditSwAssy";
            this.btnEditSwAssy.Size = new System.Drawing.Size(111, 42);
            this.btnEditSwAssy.TabIndex = 1;
            this.btnEditSwAssy.Text = "修改装配体";
            this.btnEditSwAssy.UseVisualStyleBackColor = true;
            this.btnEditSwAssy.Click += new System.EventHandler(this.btnEditSwAssy_Click);
            // 
            // FrmMian
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
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
    }
}

