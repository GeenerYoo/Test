namespace Addin_Test
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnCurrentPrjPath = new System.Windows.Forms.Button();
            this.btnProjectPropCheck = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnCurrentPrjPath
            // 
            this.btnCurrentPrjPath.Location = new System.Drawing.Point(57, 37);
            this.btnCurrentPrjPath.Name = "btnCurrentPrjPath";
            this.btnCurrentPrjPath.Size = new System.Drawing.Size(123, 30);
            this.btnCurrentPrjPath.TabIndex = 0;
            this.btnCurrentPrjPath.Text = "프로젝트 파일 경로";
            this.btnCurrentPrjPath.UseVisualStyleBackColor = true;
            this.btnCurrentPrjPath.Click += new System.EventHandler(this.btnCurrentPrjPath_Click);
            // 
            // btnProjectPropCheck
            // 
            this.btnProjectPropCheck.Location = new System.Drawing.Point(57, 105);
            this.btnProjectPropCheck.Name = "btnProjectPropCheck";
            this.btnProjectPropCheck.Size = new System.Drawing.Size(123, 30);
            this.btnProjectPropCheck.TabIndex = 1;
            this.btnProjectPropCheck.Text = "프로젝트 속성 체크";
            this.btnProjectPropCheck.UseVisualStyleBackColor = true;
            this.btnProjectPropCheck.Click += new System.EventHandler(this.btnProjectPropCheck_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(769, 477);
            this.Controls.Add(this.btnProjectPropCheck);
            this.Controls.Add(this.btnCurrentPrjPath);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnCurrentPrjPath;
        private System.Windows.Forms.Button btnProjectPropCheck;
    }
}