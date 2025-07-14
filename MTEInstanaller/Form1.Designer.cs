namespace MTEInstaller
{
    partial class Form1
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
            this.btn_install = new System.Windows.Forms.Button();
            this.btn_uninstall = new System.Windows.Forms.Button();
            this.label_version = new System.Windows.Forms.Label();
            this.btn_update = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn_install
            // 
            this.btn_install.Location = new System.Drawing.Point(12, 50);
            this.btn_install.Name = "btn_install";
            this.btn_install.Size = new System.Drawing.Size(75, 23);
            this.btn_install.TabIndex = 0;
            this.btn_install.Text = "安装";
            this.btn_install.UseVisualStyleBackColor = true;
            this.btn_install.Click += new System.EventHandler(this.btn_install_Click);
            // 
            // btn_uninstall
            // 
            this.btn_uninstall.Location = new System.Drawing.Point(93, 50);
            this.btn_uninstall.Name = "btn_uninstall";
            this.btn_uninstall.Size = new System.Drawing.Size(75, 23);
            this.btn_uninstall.TabIndex = 1;
            this.btn_uninstall.Text = "卸载";
            this.btn_uninstall.UseVisualStyleBackColor = true;
            this.btn_uninstall.Click += new System.EventHandler(this.btn_uninstall_Click);
            // 
            // label_version
            // 
            this.label_version.AutoSize = true;
            this.label_version.Location = new System.Drawing.Point(12, 18);
            this.label_version.Name = "label_version";
            this.label_version.Size = new System.Drawing.Size(47, 12);
            this.label_version.TabIndex = 2;
            this.label_version.Text = "版本号:";
            // 
            // btn_update
            // 
            this.btn_update.Location = new System.Drawing.Point(174, 50);
            this.btn_update.Name = "btn_update";
            this.btn_update.Size = new System.Drawing.Size(75, 23);
            this.btn_update.TabIndex = 3;
            this.btn_update.Text = "升级";
            this.btn_update.UseVisualStyleBackColor = true;
            this.btn_update.Click += new System.EventHandler(this.btn_update_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(276, 94);
            this.Controls.Add(this.btn_update);
            this.Controls.Add(this.label_version);
            this.Controls.Add(this.btn_uninstall);
            this.Controls.Add(this.btn_install);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "安装器";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_install;
        private System.Windows.Forms.Button btn_uninstall;
        private System.Windows.Forms.Label label_version;
        private System.Windows.Forms.Button btn_update;
    }
}

