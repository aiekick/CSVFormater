/*
 * Created by SharpDevelop.
 * User: p038564
 * Date: 6/8/2010
 * Time: 4:21 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace CSVFormater
{
	partial class MainForm
	{
		/// <summary>
		/// Designer variable used to keep track of non-visual components.
		/// </summary>
		private System.ComponentModel.IContainer components = null;
		
		/// <summary>
		/// Disposes resources used by the form.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing) {
				if (components != null) {
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}
		
		/// <summary>
		/// This method is required for Windows Forms designer support.
		/// Do not change the method contents inside the source code editor. The Forms designer might
		/// not be able to load this method if it was changed manually.
		/// </summary>
		private void InitializeComponent()
		{
			this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
			this.button1 = new System.Windows.Forms.Button();
			this.RunBtn = new System.Windows.Forms.Button();
			this.ConfigBox = new System.Windows.Forms.ComboBox();
			this.button3 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.progressCalcul1Bar = new System.Windows.Forms.ProgressBar();
			this.label1 = new System.Windows.Forms.Label();
			this.progressCalcul2Bar = new System.Windows.Forms.ProgressBar();
			this.label2 = new System.Windows.Forms.Label();
			this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
			this.PathBox = new System.Windows.Forms.TextBox();
			this.checkBox1 = new System.Windows.Forms.CheckBox();
			this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
			this.progressBarBis = new System.Windows.Forms.ProgressBar();
			this.progressBar = new System.Windows.Forms.ProgressBar();
			this.listView1 = new System.Windows.Forms.ListView();
			this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
			this.tableLayoutPanel1.SuspendLayout();
			this.tableLayoutPanel2.SuspendLayout();
			this.tableLayoutPanel3.SuspendLayout();
			this.SuspendLayout();
			// 
			// tableLayoutPanel1
			// 
			this.tableLayoutPanel1.AutoSize = true;
			this.tableLayoutPanel1.ColumnCount = 2;
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 50F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
			this.tableLayoutPanel1.Controls.Add(this.button1, 0, 0);
			this.tableLayoutPanel1.Controls.Add(this.RunBtn, 0, 2);
			this.tableLayoutPanel1.Controls.Add(this.ConfigBox, 1, 1);
			this.tableLayoutPanel1.Controls.Add(this.button3, 0, 1);
			this.tableLayoutPanel1.Controls.Add(this.button2, 0, 5);
			this.tableLayoutPanel1.Controls.Add(this.progressCalcul1Bar, 1, 3);
			this.tableLayoutPanel1.Controls.Add(this.label1, 0, 3);
			this.tableLayoutPanel1.Controls.Add(this.progressCalcul2Bar, 1, 4);
			this.tableLayoutPanel1.Controls.Add(this.label2, 0, 4);
			this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 1, 0);
			this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel3, 1, 2);
			this.tableLayoutPanel1.Controls.Add(this.listView1, 1, 5);
			this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
			this.tableLayoutPanel1.Name = "tableLayoutPanel1";
			this.tableLayoutPanel1.RowCount = 6;
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tableLayoutPanel1.Size = new System.Drawing.Size(822, 261);
			this.tableLayoutPanel1.TabIndex = 2;
			// 
			// button1
			// 
			this.button1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.button1.Location = new System.Drawing.Point(3, 3);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(44, 28);
			this.button1.TabIndex = 0;
			this.button1.Text = "...";
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Click += new System.EventHandler(this.OpenFiles_Click);
			// 
			// RunBtn
			// 
			this.RunBtn.Dock = System.Windows.Forms.DockStyle.Fill;
			this.RunBtn.Location = new System.Drawing.Point(3, 68);
			this.RunBtn.Name = "RunBtn";
			this.RunBtn.Size = new System.Drawing.Size(44, 22);
			this.RunBtn.TabIndex = 2;
			this.RunBtn.Text = "GO";
			this.RunBtn.UseVisualStyleBackColor = true;
			this.RunBtn.Click += new System.EventHandler(this.LaunchThread_Click);
			// 
			// ConfigBox
			// 
			this.ConfigBox.Dock = System.Windows.Forms.DockStyle.Fill;
			this.ConfigBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ConfigBox.FormattingEnabled = true;
			this.ConfigBox.Location = new System.Drawing.Point(53, 37);
			this.ConfigBox.Name = "ConfigBox";
			this.ConfigBox.Size = new System.Drawing.Size(767, 21);
			this.ConfigBox.TabIndex = 10;
			// 
			// button3
			// 
			this.button3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.button3.Location = new System.Drawing.Point(3, 37);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(44, 25);
			this.button3.TabIndex = 11;
			this.button3.TabStop = false;
			this.button3.Text = "Edit";
			this.button3.UseVisualStyleBackColor = true;
			this.button3.Click += new System.EventHandler(this.EditFile_Click);
			// 
			// button2
			// 
			this.button2.Dock = System.Windows.Forms.DockStyle.Top;
			this.button2.Location = new System.Drawing.Point(3, 154);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(44, 23);
			this.button2.TabIndex = 13;
			this.button2.Text = "Save";
			this.button2.UseVisualStyleBackColor = true;
			this.button2.Click += new System.EventHandler(this.Save_Click);
			// 
			// progressCalcul1Bar
			// 
			this.progressCalcul1Bar.Dock = System.Windows.Forms.DockStyle.Fill;
			this.progressCalcul1Bar.Location = new System.Drawing.Point(53, 96);
			this.progressCalcul1Bar.Name = "progressCalcul1Bar";
			this.progressCalcul1Bar.Size = new System.Drawing.Size(767, 23);
			this.progressCalcul1Bar.TabIndex = 14;
			// 
			// label1
			// 
			this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label1.Location = new System.Drawing.Point(3, 93);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 29);
			this.label1.TabIndex = 15;
			this.label1.Text = "Calcul1";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// progressCalcul2Bar
			// 
			this.progressCalcul2Bar.Dock = System.Windows.Forms.DockStyle.Fill;
			this.progressCalcul2Bar.Location = new System.Drawing.Point(53, 125);
			this.progressCalcul2Bar.Name = "progressCalcul2Bar";
			this.progressCalcul2Bar.Size = new System.Drawing.Size(767, 23);
			this.progressCalcul2Bar.TabIndex = 16;
			// 
			// label2
			// 
			this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.label2.Location = new System.Drawing.Point(3, 122);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 29);
			this.label2.TabIndex = 17;
			this.label2.Text = "Calcul2";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// tableLayoutPanel2
			// 
			this.tableLayoutPanel2.AutoSize = true;
			this.tableLayoutPanel2.ColumnCount = 2;
			this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
			this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
			this.tableLayoutPanel2.Controls.Add(this.PathBox, 0, 0);
			this.tableLayoutPanel2.Controls.Add(this.checkBox1, 1, 0);
			this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tableLayoutPanel2.Location = new System.Drawing.Point(53, 3);
			this.tableLayoutPanel2.Name = "tableLayoutPanel2";
			this.tableLayoutPanel2.RowCount = 1;
			this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tableLayoutPanel2.Size = new System.Drawing.Size(767, 28);
			this.tableLayoutPanel2.TabIndex = 18;
			// 
			// PathBox
			// 
			this.PathBox.Dock = System.Windows.Forms.DockStyle.Fill;
			this.PathBox.Location = new System.Drawing.Point(3, 3);
			this.PathBox.Name = "PathBox";
			this.PathBox.Size = new System.Drawing.Size(605, 20);
			this.PathBox.TabIndex = 2;
			// 
			// checkBox1
			// 
			this.checkBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.checkBox1.Location = new System.Drawing.Point(614, 3);
			this.checkBox1.MaximumSize = new System.Drawing.Size(0, 20);
			this.checkBox1.MinimumSize = new System.Drawing.Size(150, 0);
			this.checkBox1.Name = "checkBox1";
			this.checkBox1.Size = new System.Drawing.Size(150, 20);
			this.checkBox1.TabIndex = 3;
			this.checkBox1.Text = "Append On Current Excel";
			this.checkBox1.UseVisualStyleBackColor = true;
			// 
			// tableLayoutPanel3
			// 
			this.tableLayoutPanel3.ColumnCount = 1;
			this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
			this.tableLayoutPanel3.Controls.Add(this.progressBarBis, 0, 1);
			this.tableLayoutPanel3.Controls.Add(this.progressBar, 0, 0);
			this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tableLayoutPanel3.Location = new System.Drawing.Point(50, 65);
			this.tableLayoutPanel3.Margin = new System.Windows.Forms.Padding(0);
			this.tableLayoutPanel3.Name = "tableLayoutPanel3";
			this.tableLayoutPanel3.RowCount = 2;
			this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
			this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
			this.tableLayoutPanel3.Size = new System.Drawing.Size(773, 28);
			this.tableLayoutPanel3.TabIndex = 20;
			// 
			// progressBarBis
			// 
			this.progressBarBis.Dock = System.Windows.Forms.DockStyle.Fill;
			this.progressBarBis.Location = new System.Drawing.Point(3, 15);
			this.progressBarBis.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
			this.progressBarBis.Name = "progressBarBis";
			this.progressBarBis.Size = new System.Drawing.Size(767, 12);
			this.progressBarBis.TabIndex = 5;
			// 
			// progressBar
			// 
			this.progressBar.Dock = System.Windows.Forms.DockStyle.Fill;
			this.progressBar.Location = new System.Drawing.Point(3, 1);
			this.progressBar.Margin = new System.Windows.Forms.Padding(3, 1, 3, 1);
			this.progressBar.Name = "progressBar";
			this.progressBar.Size = new System.Drawing.Size(767, 12);
			this.progressBar.TabIndex = 4;
			// 
			// listView1
			// 
			this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
			this.columnHeader1});
			this.listView1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.listView1.Location = new System.Drawing.Point(53, 154);
			this.listView1.Name = "listView1";
			this.listView1.Size = new System.Drawing.Size(767, 147);
			this.listView1.TabIndex = 21;
			this.listView1.UseCompatibleStateImageBehavior = false;
			this.listView1.View = System.Windows.Forms.View.Details;
			// 
			// columnHeader1
			// 
			this.columnHeader1.Text = "Message";
			this.columnHeader1.Width = 500;
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(822, 261);
			this.Controls.Add(this.tableLayoutPanel1);
			this.Name = "MainForm";
			this.Text = "CSVFormater";
			this.tableLayoutPanel1.ResumeLayout(false);
			this.tableLayoutPanel1.PerformLayout();
			this.tableLayoutPanel2.ResumeLayout(false);
			this.tableLayoutPanel2.PerformLayout();
			this.tableLayoutPanel3.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button RunBtn;
        public System.Windows.Forms.ComboBox ConfigBox;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ProgressBar progressCalcul1Bar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar progressCalcul2Bar;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.TextBox PathBox;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.ProgressBar progressBarBis;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader columnHeader1;
    }
}
