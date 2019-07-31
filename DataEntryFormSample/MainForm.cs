using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataEntryFormSample {
    public partial class MainForm : DevExpress.XtraBars.Ribbon.RibbonForm {
        public MainForm() {
            InitializeComponent();
            LoadDocumentTemplate();
        }

        private void LoadDocumentTemplate() {
            spreadsheetControl1.LoadDocument("PayrollCalculator_template.xlsx");
        }
    }
}
