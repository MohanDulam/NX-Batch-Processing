
namespace NX_Batch_Processing
{
    partial class NX_Batch_Processing
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NX_Batch_Processing));
            this.label1 = new System.Windows.Forms.Label();
            this.txtInputFolder = new System.Windows.Forms.TextBox();
            this.btnBrowseInputFile = new System.Windows.Forms.Button();
            this.btnBrowseExportFolder = new System.Windows.Forms.Button();
            this.txtExportFolder = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.chkbxStep = new System.Windows.Forms.CheckBox();
            this.chkbxIGES = new System.Windows.Forms.CheckBox();
            this.chkbxParasolid = new System.Windows.Forms.CheckBox();
            this.chkbxPDF = new System.Windows.Forms.CheckBox();
            this.chkbxDXF = new System.Windows.Forms.CheckBox();
            this.chkbxDWG = new System.Windows.Forms.CheckBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.lblInfo = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Century", 11F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label1.Location = new System.Drawing.Point(11, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(216, 27);
            this.label1.TabIndex = 0;
            this.label1.Text = "Input Files Folder";
            // 
            // txtInputFolder
            // 
            this.txtInputFolder.Location = new System.Drawing.Point(244, 39);
            this.txtInputFolder.Name = "txtInputFolder";
            this.txtInputFolder.Size = new System.Drawing.Size(351, 26);
            this.txtInputFolder.TabIndex = 1;
            // 
            // btnBrowseInputFile
            // 
            this.btnBrowseInputFile.Font = new System.Drawing.Font("Century", 13F, System.Drawing.FontStyle.Bold);
            this.btnBrowseInputFile.Location = new System.Drawing.Point(608, 29);
            this.btnBrowseInputFile.Name = "btnBrowseInputFile";
            this.btnBrowseInputFile.Size = new System.Drawing.Size(148, 47);
            this.btnBrowseInputFile.TabIndex = 2;
            this.btnBrowseInputFile.Text = "Browse";
            this.btnBrowseInputFile.UseVisualStyleBackColor = true;
            this.btnBrowseInputFile.Click += new System.EventHandler(this.btnBrowseInputFile_Click);
            // 
            // btnBrowseExportFolder
            // 
            this.btnBrowseExportFolder.Font = new System.Drawing.Font("Century", 13F, System.Drawing.FontStyle.Bold);
            this.btnBrowseExportFolder.Location = new System.Drawing.Point(608, 100);
            this.btnBrowseExportFolder.Name = "btnBrowseExportFolder";
            this.btnBrowseExportFolder.Size = new System.Drawing.Size(148, 47);
            this.btnBrowseExportFolder.TabIndex = 5;
            this.btnBrowseExportFolder.Text = "Browse";
            this.btnBrowseExportFolder.UseVisualStyleBackColor = true;
            this.btnBrowseExportFolder.Click += new System.EventHandler(this.btnBrowseExportFolder_Click);
            // 
            // txtExportFolder
            // 
            this.txtExportFolder.Location = new System.Drawing.Point(244, 110);
            this.txtExportFolder.Name = "txtExportFolder";
            this.txtExportFolder.Size = new System.Drawing.Size(351, 26);
            this.txtExportFolder.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Century", 11F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label2.Location = new System.Drawing.Point(10, 107);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(169, 27);
            this.label2.TabIndex = 3;
            this.label2.Text = "Export Folder";
            // 
            // chkbxStep
            // 
            this.chkbxStep.AutoSize = true;
            this.chkbxStep.Font = new System.Drawing.Font("Century", 12F, System.Drawing.FontStyle.Bold);
            this.chkbxStep.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.chkbxStep.Location = new System.Drawing.Point(51, 164);
            this.chkbxStep.Name = "chkbxStep";
            this.chkbxStep.Size = new System.Drawing.Size(106, 32);
            this.chkbxStep.TabIndex = 6;
            this.chkbxStep.Text = "STEP";
            this.chkbxStep.UseVisualStyleBackColor = true;
            // 
            // chkbxIGES
            // 
            this.chkbxIGES.AutoSize = true;
            this.chkbxIGES.Font = new System.Drawing.Font("Century", 12F, System.Drawing.FontStyle.Bold);
            this.chkbxIGES.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.chkbxIGES.Location = new System.Drawing.Point(51, 218);
            this.chkbxIGES.Name = "chkbxIGES";
            this.chkbxIGES.Size = new System.Drawing.Size(103, 32);
            this.chkbxIGES.TabIndex = 7;
            this.chkbxIGES.Text = "IGES";
            this.chkbxIGES.UseVisualStyleBackColor = true;
            // 
            // chkbxParasolid
            // 
            this.chkbxParasolid.AutoSize = true;
            this.chkbxParasolid.Font = new System.Drawing.Font("Century", 12F, System.Drawing.FontStyle.Bold);
            this.chkbxParasolid.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.chkbxParasolid.Location = new System.Drawing.Point(51, 272);
            this.chkbxParasolid.Name = "chkbxParasolid";
            this.chkbxParasolid.Size = new System.Drawing.Size(153, 32);
            this.chkbxParasolid.TabIndex = 8;
            this.chkbxParasolid.Text = "Parasolid";
            this.chkbxParasolid.UseVisualStyleBackColor = true;
            // 
            // chkbxPDF
            // 
            this.chkbxPDF.AutoSize = true;
            this.chkbxPDF.Font = new System.Drawing.Font("Century", 12F, System.Drawing.FontStyle.Bold);
            this.chkbxPDF.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.chkbxPDF.Location = new System.Drawing.Point(248, 272);
            this.chkbxPDF.Name = "chkbxPDF";
            this.chkbxPDF.Size = new System.Drawing.Size(92, 32);
            this.chkbxPDF.TabIndex = 11;
            this.chkbxPDF.Text = "PDF";
            this.chkbxPDF.UseVisualStyleBackColor = true;
            // 
            // chkbxDXF
            // 
            this.chkbxDXF.AutoSize = true;
            this.chkbxDXF.Font = new System.Drawing.Font("Century", 12F, System.Drawing.FontStyle.Bold);
            this.chkbxDXF.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.chkbxDXF.Location = new System.Drawing.Point(248, 218);
            this.chkbxDXF.Name = "chkbxDXF";
            this.chkbxDXF.Size = new System.Drawing.Size(93, 32);
            this.chkbxDXF.TabIndex = 10;
            this.chkbxDXF.Text = "DXF";
            this.chkbxDXF.UseVisualStyleBackColor = true;
            // 
            // chkbxDWG
            // 
            this.chkbxDWG.AutoSize = true;
            this.chkbxDWG.Font = new System.Drawing.Font("Century", 12F, System.Drawing.FontStyle.Bold);
            this.chkbxDWG.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.chkbxDWG.Location = new System.Drawing.Point(248, 164);
            this.chkbxDWG.Name = "chkbxDWG";
            this.chkbxDWG.Size = new System.Drawing.Size(103, 32);
            this.chkbxDWG.TabIndex = 9;
            this.chkbxDWG.Text = "DWG";
            this.chkbxDWG.UseVisualStyleBackColor = true;
            // 
            // btnExport
            // 
            this.btnExport.Font = new System.Drawing.Font("Century", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExport.Location = new System.Drawing.Point(437, 164);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(236, 79);
            this.btnExport.TabIndex = 13;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnClear
            // 
            this.btnClear.Font = new System.Drawing.Font("Century", 13F, System.Drawing.FontStyle.Bold);
            this.btnClear.Location = new System.Drawing.Point(485, 264);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(148, 47);
            this.btnClear.TabIndex = 12;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // lblInfo
            // 
            this.lblInfo.AutoSize = true;
            this.lblInfo.Font = new System.Drawing.Font("Century", 12F, System.Drawing.FontStyle.Bold);
            this.lblInfo.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.lblInfo.Location = new System.Drawing.Point(14, 328);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(171, 28);
            this.lblInfo.TabIndex = 14;
            this.lblInfo.Text = "Status Info...";
            // 
            // NX_Batch_Processing
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(124)))), ((int)(((byte)(98)))), ((int)(((byte)(157)))));
            this.ClientSize = new System.Drawing.Size(772, 371);
            this.Controls.Add(this.lblInfo);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.chkbxPDF);
            this.Controls.Add(this.chkbxDXF);
            this.Controls.Add(this.chkbxDWG);
            this.Controls.Add(this.chkbxParasolid);
            this.Controls.Add(this.chkbxIGES);
            this.Controls.Add(this.chkbxStep);
            this.Controls.Add(this.btnBrowseExportFolder);
            this.Controls.Add(this.txtExportFolder);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnBrowseInputFile);
            this.Controls.Add(this.txtInputFolder);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "NX_Batch_Processing";
            this.Text = "NX File Export";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion


        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtInputFolder;
        private System.Windows.Forms.Button btnBrowseInputFile;
        private System.Windows.Forms.Button btnBrowseExportFolder;
        private System.Windows.Forms.TextBox txtExportFolder;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox chkbxStep;
        private System.Windows.Forms.CheckBox chkbxIGES;
        private System.Windows.Forms.CheckBox chkbxParasolid;
        private System.Windows.Forms.CheckBox chkbxPDF;
        private System.Windows.Forms.CheckBox chkbxDXF;
        private System.Windows.Forms.CheckBox chkbxDWG;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Label lblInfo;
    }
}

