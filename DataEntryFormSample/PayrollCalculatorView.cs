﻿using DevExpress.Spreadsheet;
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
            cellBindings.Add("GrossPay", "D18");
            cellBindings.Add("TaxesAndDeductions", "D20");
            cellBindings.Add("OtherDeduction", "D22");
            cellBindings.Add("NetPay", "D24");

            cellBindings.Add("TaxStatus", "I4");
            cellBindings.Add("FederalAllowance", "I6");
            cellBindings.Add("StateTax", "I8");
            cellBindings.Add("FederalIncomeTax", "I10");
            cellBindings.Add("SocialSecurityTax", "I12");
            cellBindings.Add("MedicareTax", "I14");
            cellBindings.Add("TotalTaxesWithheld", "I16");

            cellBindings.Add("InsuranceDeduction", "I20");
            cellBindings.Add("OtherRegularDeduction", "I22");
            cellBindings.Add("TotalRegularDeductions", "I24");
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
        public string EmployeeName {
            get => GetBoundCellValue().TextValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        public double HourlyWages {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
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
        public double GrossPay => GetBoundCellValue().NumericValue;

        [Browsable(false)]
        public double TaxesAndDeductions => GetBoundCellValue().NumericValue;


        [Browsable(false)]
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
        public double NetPay => GetBoundCellValue().NumericValue;


        [Browsable(false)]
        public int TaxStatus {
            get => (int)GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        public int FederalAllowance {
            get => (int)GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        public double StateTax {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        public double FederalIncomeTax {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        public double SocialSecurityTax {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        public double MedicareTax {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        public double TotalTaxesWithheld => GetBoundCellValue().NumericValue;

        [Browsable(false)]
        public double InsuranceDeduction {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
        public double OtherRegularDeduction {
            get => GetBoundCellValue().NumericValue;
            set => SetBoundCellValue(value);
        }

        [Browsable(false)]
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
            if (control != null)
                control.CellValueChanged += SpreadsheetControl_CellValueChanged;
        }

        private void UnsubscribeEvents() {
            if (control != null)
                control.CellValueChanged -= SpreadsheetControl_CellValueChanged;
        }

        private void SpreadsheetControl_CellValueChanged(object sender, SpreadsheetCellEventArgs e) {
            if (e.SheetName != "Payroll Calculator")
                return;
            string reference = e.Cell.GetReferenceA1();
            string propertyName = cellBindings.SingleOrDefault(p => p.Value == reference).Key;
            if (!string.IsNullOrEmpty(propertyName))
                NotifyPropertyChanged(propertyName);
        }

        private Worksheet Sheet => control?.Document.Worksheets["Payroll Calculator"];

        private CellValue GetCellValue(string reference) => Sheet?[reference].Value ?? CellValue.Empty;

        private void SetCellValue(string reference, CellValue value) {
            if (Sheet != null)
                Sheet[reference].Value = value;
        }

        private CellValue GetBoundCellValue([CallerMemberName] string propertyName = "") => GetCellValue(cellBindings[propertyName]);

        private void SetBoundCellValue(CellValue value, [CallerMemberName] string propertyName = "") => SetCellValue(cellBindings[propertyName], value);
    }
}