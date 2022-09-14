
namespace H5Net
{
    partial class APIControls
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.cmbProgram = new System.Windows.Forms.ComboBox();
            this.cmbTransaction = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnCreateHeader = new System.Windows.Forms.Button();
            this.btnExecTrans = new System.Windows.Forms.Button();
            this.txtResponse = new System.Windows.Forms.TextBox();
            this.btnCrtJsonFrSel = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.lblEnv = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbDivision = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // cmbProgram
            // 
            this.cmbProgram.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cmbProgram.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbProgram.ForeColor = System.Drawing.Color.Blue;
            this.cmbProgram.FormattingEnabled = true;
            this.cmbProgram.Location = new System.Drawing.Point(4, 131);
            this.cmbProgram.Margin = new System.Windows.Forms.Padding(4);
            this.cmbProgram.Name = "cmbProgram";
            this.cmbProgram.Size = new System.Drawing.Size(367, 24);
            this.cmbProgram.TabIndex = 1;
            this.cmbProgram.SelectedIndexChanged += new System.EventHandler(this.cmbProgram_SelectedIndexChanged);
            // 
            // cmbTransaction
            // 
            this.cmbTransaction.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cmbTransaction.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbTransaction.ForeColor = System.Drawing.Color.Blue;
            this.cmbTransaction.FormattingEnabled = true;
            this.cmbTransaction.Location = new System.Drawing.Point(4, 191);
            this.cmbTransaction.Margin = new System.Windows.Forms.Padding(4);
            this.cmbTransaction.Name = "cmbTransaction";
            this.cmbTransaction.Size = new System.Drawing.Size(367, 24);
            this.cmbTransaction.TabIndex = 3;
            this.cmbTransaction.SelectedIndexChanged += new System.EventHandler(this.cmbTransaction_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 106);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 16);
            this.label1.TabIndex = 8;
            this.label1.Text = "Program";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 166);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(81, 16);
            this.label2.TabIndex = 9;
            this.label2.Text = "Transaction";
            // 
            // btnCreateHeader
            // 
            this.btnCreateHeader.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCreateHeader.Location = new System.Drawing.Point(4, 248);
            this.btnCreateHeader.Margin = new System.Windows.Forms.Padding(4);
            this.btnCreateHeader.Name = "btnCreateHeader";
            this.btnCreateHeader.Size = new System.Drawing.Size(135, 44);
            this.btnCreateHeader.TabIndex = 13;
            this.btnCreateHeader.Text = "Create Header";
            this.btnCreateHeader.UseVisualStyleBackColor = true;
            this.btnCreateHeader.Click += new System.EventHandler(this.btnCreateHeader_Click);
            // 
            // btnExecTrans
            // 
            this.btnExecTrans.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnExecTrans.Location = new System.Drawing.Point(4, 672);
            this.btnExecTrans.Margin = new System.Windows.Forms.Padding(4);
            this.btnExecTrans.Name = "btnExecTrans";
            this.btnExecTrans.Size = new System.Drawing.Size(135, 44);
            this.btnExecTrans.TabIndex = 16;
            this.btnExecTrans.Text = "Exec Transaction";
            this.btnExecTrans.UseVisualStyleBackColor = true;
            this.btnExecTrans.Click += new System.EventHandler(this.btnExecTrans_Click);
            // 
            // txtResponse
            // 
            this.txtResponse.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtResponse.CausesValidation = false;
            this.txtResponse.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtResponse.Location = new System.Drawing.Point(4, 301);
            this.txtResponse.Multiline = true;
            this.txtResponse.Name = "txtResponse";
            this.txtResponse.ReadOnly = true;
            this.txtResponse.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtResponse.Size = new System.Drawing.Size(367, 364);
            this.txtResponse.TabIndex = 17;
            // 
            // btnCrtJsonFrSel
            // 
            this.btnCrtJsonFrSel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCrtJsonFrSel.Location = new System.Drawing.Point(231, 248);
            this.btnCrtJsonFrSel.Margin = new System.Windows.Forms.Padding(4);
            this.btnCrtJsonFrSel.Name = "btnCrtJsonFrSel";
            this.btnCrtJsonFrSel.Size = new System.Drawing.Size(140, 44);
            this.btnCrtJsonFrSel.TabIndex = 18;
            this.btnCrtJsonFrSel.Text = "Create JSON";
            this.btnCrtJsonFrSel.UseVisualStyleBackColor = true;
            this.btnCrtJsonFrSel.Click += new System.EventHandler(this.btnCrtJsonFrSel_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 16);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(86, 16);
            this.label3.TabIndex = 19;
            this.label3.Text = "Environment";
            // 
            // lblEnv
            // 
            this.lblEnv.AutoSize = true;
            this.lblEnv.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.lblEnv.Location = new System.Drawing.Point(100, 16);
            this.lblEnv.Name = "lblEnv";
            this.lblEnv.Size = new System.Drawing.Size(0, 16);
            this.lblEnv.TabIndex = 20;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(4, 46);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 16);
            this.label4.TabIndex = 22;
            this.label4.Text = "Division";
            // 
            // cmbDivision
            // 
            this.cmbDivision.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cmbDivision.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbDivision.ForeColor = System.Drawing.Color.Blue;
            this.cmbDivision.FormattingEnabled = true;
            this.cmbDivision.Location = new System.Drawing.Point(4, 71);
            this.cmbDivision.Margin = new System.Windows.Forms.Padding(4);
            this.cmbDivision.Name = "cmbDivision";
            this.cmbDivision.Size = new System.Drawing.Size(367, 24);
            this.cmbDivision.TabIndex = 21;
            this.cmbDivision.SelectedIndexChanged += new System.EventHandler(this.cmbDivision_SelectedIndexChanged);
            // 
            // APIControls
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cmbDivision);
            this.Controls.Add(this.lblEnv);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnCrtJsonFrSel);
            this.Controls.Add(this.txtResponse);
            this.Controls.Add(this.btnExecTrans);
            this.Controls.Add(this.btnCreateHeader);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbTransaction);
            this.Controls.Add(this.cmbProgram);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "APIControls";
            this.Size = new System.Drawing.Size(388, 744);
            this.Load += new System.EventHandler(this.APIControls_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ComboBox cmbProgram;
        private System.Windows.Forms.ComboBox cmbTransaction;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnCreateHeader;
        private System.Windows.Forms.Button btnExecTrans;
        private System.Windows.Forms.TextBox txtResponse;
        private System.Windows.Forms.Button btnCrtJsonFrSel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblEnv;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbDivision;
    }
}
