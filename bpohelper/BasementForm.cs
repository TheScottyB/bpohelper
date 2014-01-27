using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace bpohelper
{
    public partial class BasementForm : Form
    {
        private Form1 form;

        public BasementForm()
        {
            InitializeComponent();
        }

        public BasementForm(Form1 f)
        {
            form = f;
            InitializeComponent();

        }

        private void Basementform_FormClosed(object sender, FormClosedEventArgs e)
        {
            
            // basementTypeCheckedListBox.SelectedItems;
        }

        private void BasementForm_Load(object sender, EventArgs e)
        {

        }

        private void basementTypeCheckedListBox_ItemCheck(object sender, ItemCheckEventArgs e)
        {

            string basementType = basementTypeCheckedListBox.SelectedItem.ToString();

            foreach (string s in basementTypeCheckedListBox.CheckedItems)
            {
                basementType = basementType.Insert(0, s + ",");
            }

            form.SubjectBasementType = basementType;
             
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            string basementDetails = checkedListBox1.SelectedItem.ToString();

            foreach (string s in checkedListBox1.SelectedItems)
            {
                basementDetails = basementDetails.Insert(0, s + ",");
            }

            form.SubjectBasementDetails = basementDetails;
             
        }
    }
}
