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
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraSpreadsheet;

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
            sheet.CustomCellInplaceEditors.Add(sheet["D8"], CustomCellInplaceEditorType.Custom, "RegularHoursWorked");
            sheet.CustomCellInplaceEditors.Add(sheet["D10"], CustomCellInplaceEditorType.Custom, "VacationHours");
            sheet.CustomCellInplaceEditors.Add(sheet["D12"], CustomCellInplaceEditorType.Custom, "SickHours");
            sheet.CustomCellInplaceEditors.Add(sheet["D14"], CustomCellInplaceEditorType.Custom, "OvertimeHours");
            sheet.CustomCellInplaceEditors.Add(sheet["D16"], CustomCellInplaceEditorType.Custom, "OvertimeRate");
            sheet.CustomCellInplaceEditors.Add(sheet["D22"], CustomCellInplaceEditorType.Custom, "OtherDeduction");
        }

        private RepositoryItem CreateCustomEditor(string tag) {
            switch (tag) {
                case "RegularHoursWorked": return CreateSpinEdit(0, 184, 1);
                case "VacationHours":      return CreateSpinEdit(0, 184, 1);
                case "SickHours":          return CreateSpinEdit(0, 184, 1);
                case "OvertimeHours":      return CreateSpinEdit(0, 100, 1);
                case "OvertimeRate":       return CreateSpinEdit(0, 50, 1);
                case "OtherDeduction":     return CreateSpinEdit(0, 100, 1);
                default:                   return null;
            }
        }

        private RepositoryItemSpinEdit CreateSpinEdit(int minValue, int maxValue, int increment) => new RepositoryItemSpinEdit {
            AutoHeight = false,
            BorderStyle = BorderStyles.NoBorder,
            MinValue = minValue,
            MaxValue = maxValue,
            Increment = increment,
            IsFloatValue = false
        };

        private void ActivateCustomEditor() {
            var sheet = spreadsheetControl1.ActiveWorksheet;
            var editors = sheet.CustomCellInplaceEditors.GetCustomCellInplaceEditors(sheet.Selection);
            if (editors.Count == 1)
                spreadsheetControl1.OpenCellEditor(CellEditorMode.Edit);
        }

        private void spreadsheetControl1_CustomCellEdit(object sender, SpreadsheetCustomCellEditEventArgs e) {
            if (e.ValueObject.IsText)
                e.RepositoryItem = CreateCustomEditor(e.ValueObject.TextValue);
        }

        private void spreadsheetControl1_SelectionChanged(object sender, EventArgs e) => ActivateCustomEditor();

        private void spreadsheetControl1_ProtectionWarning(object sender, HandledEventArgs e) => e.Handled = true;
    }
}
