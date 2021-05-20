
namespace ToolConvertSqlToMongoDB
{
    partial class Eazy
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
            this.btnExchangeData = new System.Windows.Forms.Button();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.txtDatabase2 = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.btnTestConnect2 = new System.Windows.Forms.Button();
            this.txtPassword2 = new System.Windows.Forms.TextBox();
            this.txtLogin2 = new System.Windows.Forms.TextBox();
            this.txtServer2 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtPort = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtDatabase1 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.btnTestConnect1 = new System.Windows.Forms.Button();
            this.txtPassword1 = new System.Windows.Forms.TextBox();
            this.txtLogin1 = new System.Windows.Forms.TextBox();
            this.txtServer1 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 16);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Location = new System.Drawing.Point(0, 311);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(504, 22);
            this.statusStrip1.TabIndex = 17;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // txtDatabase2
            // 
            this.txtDatabase2.Location = new System.Drawing.Point(82, 141);
            this.txtDatabase2.Name = "txtDatabase2";
            this.txtDatabase2.Size = new System.Drawing.Size(127, 20);
            this.txtDatabase2.TabIndex = 3;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.txtDatabase2);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.btnTestConnect2);
            this.groupBox2.Controls.Add(this.txtPassword2);
            this.groupBox2.Controls.Add(this.txtLogin2);
            this.groupBox2.Controls.Add(this.btnExchangeData);
            this.groupBox2.Controls.Add(this.txtServer2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Location = new System.Drawing.Point(258, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(234, 247);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Kết nối 2 (Dữ liệu cũ)";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(20, 144);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(56, 13);
            this.label8.TabIndex = 4;
            this.label8.Text = "Database:";
            // 
            // btnTestConnect2
            // 
            this.btnTestConnect2.Location = new System.Drawing.Point(134, 173);
            this.btnTestConnect2.Name = "btnTestConnect2";
            this.btnTestConnect2.Size = new System.Drawing.Size(75, 23);
            this.btnTestConnect2.TabIndex = 4;
            this.btnTestConnect2.Text = "Test kết nối";
            this.btnTestConnect2.UseVisualStyleBackColor = true;
            this.btnTestConnect2.Click += new System.EventHandler(this.btnTestConnect2_Click);
            // 
            // txtPassword2
            // 
            this.txtPassword2.Location = new System.Drawing.Point(82, 104);
            this.txtPassword2.Name = "txtPassword2";
            this.txtPassword2.PasswordChar = '*';
            this.txtPassword2.Size = new System.Drawing.Size(127, 20);
            this.txtPassword2.TabIndex = 2;
            // 
            // txtLogin2
            // 
            this.txtLogin2.Location = new System.Drawing.Point(82, 65);
            this.txtLogin2.Name = "txtLogin2";
            this.txtLogin2.Size = new System.Drawing.Size(127, 20);
            this.txtLogin2.TabIndex = 1;
            // 
            // txtServer2
            // 
            this.txtServer2.Location = new System.Drawing.Point(82, 28);
            this.txtServer2.Name = "txtServer2";
            this.txtServer2.Size = new System.Drawing.Size(127, 20);
            this.txtServer2.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 107);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Password:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(20, 68);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(36, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "Login:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(20, 31);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(41, 13);
            this.label6.TabIndex = 0;
            this.label6.Text = "Server:";
            // 
            // btnExchangeData
            // 
            this.btnExchangeData.Enabled = false;
            this.btnExchangeData.Location = new System.Drawing.Point(23, 211);
            this.btnExchangeData.Name = "btnExchangeData";
            this.btnExchangeData.Size = new System.Drawing.Size(118, 23);
            this.btnExchangeData.TabIndex = 10;
            this.btnExchangeData.Text = "Chuyển đổi dữ liệu";
            this.btnExchangeData.UseVisualStyleBackColor = true;
            this.btnExchangeData.Visible = true;
            this.btnExchangeData.Click += new System.EventHandler(this.btnExchangeData_Click);

            // 
           
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 31);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Server:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.txtPort);
            this.groupBox1.Controls.Add(this.txtDatabase1);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.btnTestConnect1);
            this.groupBox1.Controls.Add(this.txtPassword1);
            this.groupBox1.Controls.Add(this.txtLogin1);
            this.groupBox1.Controls.Add(this.txtServer1);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(236, 247);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Kết nối 1 (Eazy mới)";
            // 
         
           
            // 
            // txtDatabase1
            // 
            this.txtDatabase1.Location = new System.Drawing.Point(82, 141);
            this.txtDatabase1.Name = "txtDatabase1";
            this.txtDatabase1.Size = new System.Drawing.Size(127, 20);
            this.txtDatabase1.TabIndex = 3;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(20, 144);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(56, 13);
            this.label7.TabIndex = 2;
            this.label7.Text = "Database:";
            // 
            // btnTestConnect1
            // 
            this.btnTestConnect1.Location = new System.Drawing.Point(134, 211);
            this.btnTestConnect1.Name = "btnTestConnect1";
            this.btnTestConnect1.Size = new System.Drawing.Size(75, 23);
            this.btnTestConnect1.TabIndex = 4;
            this.btnTestConnect1.Text = "Test kết nối";
            this.btnTestConnect1.UseVisualStyleBackColor = true;
            this.btnTestConnect1.Click += new System.EventHandler(this.btnTestConnect1_Click);
            // 
            // txtPassword1
            // 
            this.txtPassword1.Location = new System.Drawing.Point(82, 104);
            this.txtPassword1.Name = "txtPassword1";
            this.txtPassword1.PasswordChar = '*';
            this.txtPassword1.Size = new System.Drawing.Size(127, 20);
            this.txtPassword1.TabIndex = 2;
            // 
            // txtLogin1
            // 
            this.txtLogin1.Location = new System.Drawing.Point(82, 65);
            this.txtLogin1.Name = "txtLogin1";
            this.txtLogin1.Size = new System.Drawing.Size(127, 20);
            this.txtLogin1.TabIndex = 1;
            // 
            // txtServer1
            // 
            this.txtServer1.Location = new System.Drawing.Point(82, 28);
            this.txtServer1.Name = "txtServer1";
            this.txtServer1.Size = new System.Drawing.Size(127, 20);
            this.txtServer1.TabIndex = 0;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(20, 107);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Password:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(20, 68);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(60, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "UserName:";
            // 
            // IBMS
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(504, 333);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.Name = "Eazy";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Chuyển đổi dữ liệu Eazy";
            this.Load += new System.EventHandler(this.Eazy_Load);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.TextBox txtDatabase2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnTestConnect2;
        private System.Windows.Forms.TextBox txtPassword2;
        private System.Windows.Forms.TextBox txtPort;
        private System.Windows.Forms.TextBox txtLogin2;
        private System.Windows.Forms.TextBox txtServer2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtDatabase1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnTestConnect1;
        private System.Windows.Forms.TextBox txtPassword1;
        private System.Windows.Forms.TextBox txtLogin1;
        private System.Windows.Forms.TextBox txtServer1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnExchangeData;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}