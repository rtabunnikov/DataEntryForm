using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataEntryFormSample {
    public partial class PayrollCalculatorView : Component, IBindableComponent, INotifyPropertyChanged {
        private const string payrollCalculatorSheetName = "Payroll Calculator";
        private SpreadsheetControl control;
        private ControlBindingsCollection dataBindings;
        private BindingContext bindingContext;
        private Dictionary<string, string> cellBindings = new Dictionary<string, string>();

        public PayrollCalculatorView() {
            InitializeComponent();
            CreateCellBindings();
        }

        public PayrollCalculatorView(IContainer container) {
            container.Add(this);

            InitializeComponent();
            CreateCellBindings();
        }

        private void CreateCellBindings() {
            cellBindings.Add("EmployeeName", "D4");
            cellBindings.Add("HourlyWages", "D6");
            cellBindings.Add("RegularHoursWorked", "D8");
            cellBindings.Add("VacationHours", "D10");
            cellBindings.Add("SickHours", "D12");
            cellBindings.Add("OvertimeHours", "D14");
            cellBindings.Add("OvertimeRate", "D16");
            cellBindings.Add("OtherDeduction", "D22");

            cellBindings.Add("TaxStatus", "I4");
            cellBindings.Add("FederalAllowance", "I6");
            cellBindings.Add("StateTax", "I8");
            cellBindings.Add("FederalIncomeTax", "I10");
            cellBindings.Add("SocialSecurityTax", "I12");
            cellBindings.Add("MedicareTax", "I14");

            cellBindings.Add("InsuranceDeduction", "I20");
            cellBindings.Add("OtherRegularDeduction", "I22");
        }

        #region Properties
        public SpreadsheetControl Control {
            get => control;
            set {
                if (!ReferenceEquals(control, value)) {
                    UnsubscribeEvents();
                    control = value;
                    SubscribeEvents();
                }
            }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string EmployeeName {
            get => GetBoundCellValue().TextValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double HourlyWages {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double RegularHoursWorked {
            get => GetBoundCellValue().NumericValue;
            set {
                if (GetBoundCellValue().NumericValue != value) {
                    SetBoundCellValue(value);
                    NotifyPropertyChanged();
                }
            }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double VacationHours {
            get => GetBoundCellValue().NumericValue;
            set {
                if (GetBoundCellValue().NumericValue != value) {
                    SetBoundCellValue(value);
                    NotifyPropertyChanged();
                }
            }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double SickHours {
            get => GetBoundCellValue().NumericValue;
            set {
                if (GetBoundCellValue().NumericValue != value) {
                    SetBoundCellValue(value);
                    NotifyPropertyChanged();
                }
            }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double OvertimeHours {
            get => GetBoundCellValue().NumericValue;
            set {
                if (GetBoundCellValue().NumericValue != value) {
                    SetBoundCellValue(value);
                    NotifyPropertyChanged();
                }
            }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double OvertimeRate {
            get => GetBoundCellValue().NumericValue;
            set {
                if (GetBoundCellValue().NumericValue != value) {
                    SetBoundCellValue(value);
                    NotifyPropertyChanged();
                }
            }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double GrossPay => GetBoundCellValue().NumericValue;

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double TaxesAndDeductions => GetBoundCellValue().NumericValue;


        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double OtherDeduction {
            get => GetBoundCellValue().NumericValue;
            set {
                if (GetBoundCellValue().NumericValue != value) {
                    SetBoundCellValue(value);
                    NotifyPropertyChanged();
                }
            }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double NetPay => GetBoundCellValue().NumericValue;


        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int TaxStatus {
            get => (int)GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int FederalAllowance {
            get => (int)GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double StateTax {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double FederalIncomeTax {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double SocialSecurityTax {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double MedicareTax {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double TotalTaxesWithheld => GetBoundCellValue().NumericValue;

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double InsuranceDeduction {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double OtherRegularDeduction {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double TotalRegularDeductions => GetBoundCellValue().NumericValue;
        #endregion

        #region IBindableComponent members
        public ControlBindingsCollection DataBindings {
            get {
                if (dataBindings == null)
                    dataBindings = new ControlBindingsCollection(this);
                return dataBindings;
            }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public BindingContext BindingContext {
            get {
                if (bindingContext == null)
                    bindingContext = new BindingContext();
                return bindingContext;
            }
            set => bindingContext = value;
        }
        #endregion

        #region INotifyPropertyChanged members
        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        private void SubscribeEvents() {
            if (control != null) {
                control.CellValueChanged += SpreadsheetControl_CellValueChanged;
                control.SelectionChanged += Control_SelectionChanged;
            }
        }

        private void UnsubscribeEvents() {
            if (control != null) {
                control.CellValueChanged -= SpreadsheetControl_CellValueChanged;
                control.SelectionChanged += Control_SelectionChanged;
            }
        }

        private void SpreadsheetControl_CellValueChanged(object sender, SpreadsheetCellEventArgs e) {
            if (e.SheetName == payrollCalculatorSheetName) {
                string reference = e.Cell.GetReferenceA1();
                string propertyName = cellBindings.SingleOrDefault(p => p.Value == reference).Key;
                if (!string.IsNullOrEmpty(propertyName))
                    NotifyPropertyChanged(propertyName);
            }
        }

        private void Control_SelectionChanged(object sender, EventArgs e) => ActivateCellEditor();

        private Worksheet Sheet => (control != null && control.Document.Worksheets.Contains(payrollCalculatorSheetName)) ?
                    control.Document.Worksheets[payrollCalculatorSheetName] : null;

        private CellValue GetCellValue(string reference) => Sheet?[reference].Value ?? CellValue.Empty;

        private void SetCellValue(string reference, CellValue value) {
            if (Sheet != null) {
                if (reference == Sheet.Selection.GetReferenceA1())
                    DeactivateCellEditor();
                Sheet[reference].Value = value;
                if (reference == Sheet.Selection.GetReferenceA1())
                    ActivateCellEditor();
            }
        }

        private CellValue GetBoundCellValue([CallerMemberName] string propertyName = "") => GetCellValue(cellBindings[propertyName]);

        private void SetBoundCellValue(CellValue value, [CallerMemberName] string propertyName = "") => SetCellValue(cellBindings[propertyName], value);

        private void ActivateCellEditor() {
            var sheet = Sheet;
            if (sheet != null) {
                var editors = sheet.CustomCellInplaceEditors.GetCustomCellInplaceEditors(sheet.Selection);
                if (editors.Count == 1)
                    control.OpenCellEditor(CellEditorMode.Edit);
            }
        }

        private void DeactivateCellEditor() {
            if (control != null && control.IsCellEditorActive)
                control.CloseCellEditor(CellEditorEnterValueMode.Cancel);
        }
    }
}
