using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TestSheetAddOn
{
    public partial class frmTDIName : Form
    {
        public frmTDIName()
        {
            InitializeComponent();
        }

        public String NameFormat;

        private void btnOK_Click(object sender, EventArgs e)
        {
            NameFormat = txtNameFormat.Text.Trim();

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {

        }
    }
}
