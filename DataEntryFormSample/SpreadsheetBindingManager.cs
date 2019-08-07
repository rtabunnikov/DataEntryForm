using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataEntryFormSample {
    public partial class SpreadsheetBindingManager : Component {
        private SpreadsheetControl control;
        private object dataSource;
        private object currentItem;
        private BindingManagerBase bindingManager;
        private readonly Dictionary<string, string> cellBindings = new Dictionary<string, string>();
        private readonly PropertyDescriptorCollection propertyDescriptors = new PropertyDescriptorCollection(null);

        public SpreadsheetBindingManager() {
            InitializeComponent();
        }

        public SpreadsheetBindingManager(IContainer container) {
            container.Add(this);
            InitializeComponent();
        }

        public SpreadsheetControl Control {
            get => control;
            set {
                if (!ReferenceEquals(control, value)) {
                    if (control != null)
                        control.CellValueChanged -= SpreadsheetControl_CellValueChanged;
                    control = value;
                    if (control != null)
                        control.CellValueChanged += SpreadsheetControl_CellValueChanged;
                }
            }
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string SheetName { get; set; }

        public object DataSource {
            get => dataSource;
            set {
                if (!ReferenceEquals(dataSource, value)) {
                    Detach();
                    dataSource = value;
                    Attach();
                }
            }
        }

        public void AddBinding(string propertyName, string cellReference) {
            if (cellBindings.ContainsKey(propertyName))
                throw new ArgumentException($"Already has binding to {propertyName} property");
            ITypedList typedList = dataSource as ITypedList;
            if (typedList != null) {
                PropertyDescriptorCollection dataSourceProperties = typedList.GetItemProperties(null);
                PropertyDescriptor propertyDescriptor = dataSourceProperties[propertyName];
                if (propertyDescriptor == null)
                    throw new InvalidOperationException($"Unknown { propertyName } property");
                if (currentItem != null)
                    propertyDescriptor.AddValueChanged(currentItem, OnPropertyChanged);
                propertyDescriptors.Add(propertyDescriptor);
            }
            cellBindings.Add(propertyName, cellReference);
        }

        public void RemoveBinding(string propertyName) {
            if (cellBindings.ContainsKey(propertyName)) {
                PropertyDescriptor propertyDescriptor = propertyDescriptors[propertyName];
                if (currentItem != null)
                    propertyDescriptor.RemoveValueChanged(currentItem, OnPropertyChanged);
                propertyDescriptors.Remove(propertyDescriptor);
                cellBindings.Remove(propertyName);
            }
        }

        public void ClearBindings() {
            UnsubscribePropertyChanged();
            propertyDescriptors.Clear();
            cellBindings.Clear();
        }

        private void Attach() {
            ICurrencyManagerProvider provider = dataSource as ICurrencyManagerProvider;
            if (provider != null) {
                bindingManager = provider.CurrencyManager;
                bindingManager.CurrentChanged += BindingManager_CurrentChanged;
                currentItem = bindingManager.Current;
            }
            ITypedList typedList = dataSource as ITypedList;
            if (typedList != null) {
                PropertyDescriptorCollection dataSourceProperties = typedList.GetItemProperties(null);
                foreach(string propertyName in cellBindings.Keys) {
                    PropertyDescriptor propertyDescriptor = dataSourceProperties[propertyName];
                    if (propertyDescriptor == null)
                        throw new InvalidOperationException($"Unable to get property descriptor for { propertyName } property");
                    propertyDescriptors.Add(propertyDescriptor);
                }
            }
            SubscribePropertyChanged();
        }

        private void Detach() {
            if (dataSource != null) {
                UnsubscribePropertyChanged();
                if (bindingManager != null) {
                    bindingManager.CurrentChanged -= BindingManager_CurrentChanged;
                    bindingManager = null;
                    currentItem = null;
                }
                propertyDescriptors.Clear();
            }
        }

        private void BindingManager_CurrentChanged(object sender, EventArgs e) {
            DeactivateCellEditor(CellEditorEnterValueMode.ActiveCell);
            control.BeginUpdate();
            try {
                UnsubscribePropertyChanged();
                currentItem = bindingManager.Current;
                PullData();
                SubscribePropertyChanged();
            }
            finally {
                control.EndUpdate();
                ActivateCellEditor();
            }
        }

        private void UnsubscribePropertyChanged() {
            if (currentItem != null) {
                foreach (PropertyDescriptor propertyDescriptor in propertyDescriptors)
                    propertyDescriptor.RemoveValueChanged(currentItem, OnPropertyChanged);
            }
        }
        private void SubscribePropertyChanged() {
            if (currentItem != null) {
                foreach (PropertyDescriptor propertyDescriptor in propertyDescriptors)
                    propertyDescriptor.AddValueChanged(currentItem, OnPropertyChanged);
            }
        }

        private void OnPropertyChanged(object sender, EventArgs eventArgs) {
            PropertyDescriptor propertyDescriptor = sender as PropertyDescriptor;
            if (propertyDescriptor != null && bindingManager != null) {
                string reference;
                if (cellBindings.TryGetValue(propertyDescriptor.Name, out reference))
                    SetCellValue(reference, CellValue.FromObject(propertyDescriptor.GetValue(currentItem)));
            }
        }

        private void PullData() {
            if (currentItem != null) {
                foreach (PropertyDescriptor propertyDescriptor in propertyDescriptors)
                    SetCellValue(cellBindings[propertyDescriptor.Name], CellValue.FromObject(propertyDescriptor.GetValue(currentItem)));
            }
        }

        private void SpreadsheetControl_CellValueChanged(object sender, SpreadsheetCellEventArgs e) {
            if (e.SheetName == SheetName) {
                string reference = e.Cell.GetReferenceA1();
                string propertyName = cellBindings.SingleOrDefault(p => p.Value == reference).Key;
                if (!string.IsNullOrEmpty(propertyName)) {
                    PropertyDescriptor propertyDescriptor = propertyDescriptors[propertyName];
                    if (propertyDescriptor != null && currentItem != null)
                        propertyDescriptor.SetValue(currentItem, e.Value.ToObject());
                }
            }
        }

        private Worksheet Sheet => 
            control != null && control.Document.Worksheets.Contains(SheetName) ? control.Document.Worksheets[SheetName] : null;

        private void SetCellValue(string reference, CellValue value) {
            if (Sheet != null) {
                if (reference == Sheet.Selection.GetReferenceA1())
                    DeactivateCellEditor();
                Sheet[reference].Value = value;
                if (reference == Sheet.Selection.GetReferenceA1())
                    ActivateCellEditor();
            }
        }

        private void ActivateCellEditor() {
            var sheet = Sheet;
            if (sheet != null) {
                var editors = sheet.CustomCellInplaceEditors.GetCustomCellInplaceEditors(sheet.Selection);
                if (editors.Count == 1)
                    control.OpenCellEditor(CellEditorMode.Edit);
            }
        }

        private void DeactivateCellEditor(CellEditorEnterValueMode mode = CellEditorEnterValueMode.Cancel) {
            if (control != null && control.IsCellEditorActive)
                control.CloseCellEditor(mode);
        }
    }
}
