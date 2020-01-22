using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SendSheetsViaMail
{
    public partial class CSVInput : Form
    {
        public CSVInput()
        {
            InitializeComponent();
        }

        DataTable dt = new DataTable("CSVInput");
        private void CSVInput_Load(object sender, EventArgs e)
        {
            PersistData data = PersistData.Load();

            if (data != null)
            {
                txtBody.Text = data.MailBody;
                dt = data.CSVConfig;
                txtSubject.Text = data.MailTitle;
            }
            else
            {
                dt.Clear();
                dt.Columns.AddRange(new DataColumn[] { new DataColumn("Department"),
               new DataColumn( "Slides"),
               new DataColumn("E-Mails")
            });
            }
            BindingSource SBind = new BindingSource();
            SBind.DataSource = dt;
            grdData.Columns.Clear();
            grdData.DataSource = SBind;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            PersistData.Save(new PersistData {MailTitle = txtSubject.Text, MailBody = txtBody.Text, CSVConfig = dt });
        }
    }
}
