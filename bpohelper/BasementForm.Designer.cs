namespace bpohelper
{
    partial class BasementForm
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
            this.basementTypeCheckedListBox = new System.Windows.Forms.CheckedListBox();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // basementTypeCheckedListBox
            // 
            this.basementTypeCheckedListBox.FormattingEnabled = true;
            this.basementTypeCheckedListBox.Items.AddRange(new object[] {
            "Full",
            "Partial",
            "Walkout",
            "English",
            "None"});
            this.basementTypeCheckedListBox.Location = new System.Drawing.Point(12, 12);
            this.basementTypeCheckedListBox.Name = "basementTypeCheckedListBox";
            this.basementTypeCheckedListBox.Size = new System.Drawing.Size(75, 79);
            this.basementTypeCheckedListBox.TabIndex = 0;
            this.basementTypeCheckedListBox.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.basementTypeCheckedListBox_ItemCheck);
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Items.AddRange(new object[] {
            "Finish",
            "Partially Finished",
            "Unfinished",
            "Crawl",
            "Cellar",
            "Sub-Basement",
            "Slab",
            "Exterior Access",
            "Other",
            "Rough‐In",
            "None"});
            this.checkedListBox1.Location = new System.Drawing.Point(142, 12);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(120, 169);
            this.checkedListBox1.TabIndex = 1;
            this.checkedListBox1.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkedListBox1_ItemCheck);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(47, 214);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "label1";
            // 
            // BasementForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.basementTypeCheckedListBox);
            this.Name = "BasementForm";
            this.Text = "Basement Details";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Basementform_FormClosed);
            this.Load += new System.EventHandler(this.BasementForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox basementTypeCheckedListBox;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Label label1;
    }
}