using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.XtraEditors.Repository;

namespace DataEntryFormSample {
    public partial class MainForm : DevExpress.XtraBars.Ribbon.RibbonForm {
        public MainForm() {
            InitializeComponent();
            LoadDocumentTemplate();
            BindCustomEditors();
        }

        private void LoadDocumentTemplate() {
            spreadsheetControl1.LoadDocument("PayrollCalculator_template.xlsx");
        }

        private void BindCustomEditors() {
            var sheet = spreadsheetControl1.ActiveWorksheet;
            var customEditors = sheet.CustomCellInplaceEditors;
            customEditors.Add(sheet["D8"], CustomCellInplaceEditorType.Custom, "RegularHoursWorked");
            customEditors.Add(sheet["D10"], CustomCellInplaceEditorType.Custom, "VacationHours");
            customEditors.Add(sheet["D12"], CustomCellInplaceEditorType.Custom, "SickHours");
            customEditors.Add(sheet["D14"], CustomCellInplaceEditorType.Custom, "OvertimeHours");
            customEditors.Add(sheet["D16"], CustomCellInplaceEditorType.Custom, "OvertimeRate");
            customEditors.Add(sheet["D22"], CustomCellInplaceEditorType.Custom, "OtherDeduction");
        }

        private void spreadsheetControl1_CustomCellEdit(object sender, DevExpress.XtraSpreadsheet.SpreadsheetCustomCellEditEventArgs e) {
            if (e.ValueObject.IsText)
                e.RepositoryItem = CreateCustomEditor(e.ValueObject.TextValue);
            if (e.RepositoryItem != null)
                e.RepositoryItem.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Default;
        }

        private RepositoryItem CreateCustomEditor(string tag) {
            switch (tag) {
                case "RegularHoursWorked":
                    return CreateSpinEdit(0, 184, 1);
                case "VacationHours":
                    return CreateSpinEdit(0, 184, 1);
                case "SickHours":
                    return CreateSpinEdit(0, 184, 1);
                case "OvertimeHours":
                    return CreateSpinEdit(0, 100, 1);
                case "OvertimeRate":
                    return CreateSpinEdit(0, 100, 1);
                case "OtherDeduction":
                    return CreateSpinEdit(0, 100, 1);
                default:
                    return null;
            }
        }

        private RepositoryItemSpinEdit CreateSpinEdit(int minValue, int maxValue, int increment) => new RepositoryItemSpinEdit {
            AutoHeight = false,
            BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder,
            MinValue = minValue,
            MaxValue = maxValue,
            Increment = increment,
            IsFloatValue = false
        };

        private void ActivateEditor() {
            var sheet = spreadsheetControl1.ActiveWorksheet;
            var editors = sheet.CustomCellInplaceEditors.GetCustomCellInplaceEditors(sheet.Selection);
            if (editors.Count == 1)
                spreadsheetControl1.OpenCellEditor(DevExpress.XtraSpreadsheet.CellEditorMode.Edit);
        }

        private void spreadsheetControl1_SelectionChanged(object sender, EventArgs e) => ActivateEditor();

        private void spreadsheetControl1_ProtectionWarning(object sender, HandledEventArgs e) => e.Handled = true;
    }
}
