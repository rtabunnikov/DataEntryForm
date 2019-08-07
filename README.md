# Spreadsheet Data Entry Form example

Platform: WinForms  
User story: payroll calculator

This example demonstrate ability to use Spreadsheet Control as data entry form.

Step by step:  
* Create document template (PayrollCalculatorTemplate.xlsx), apply worksheet protection (password is 123), protect workbook structure
* Create PayrollModel.cs - entity class, implement INotifyPropertyChanged, add properties (EmployeeName, RegularHoursWorked, etc.)
* Add spreadsheet control to main form
* Add code which load document template, bind custom cell inplace editors, create custom editors at the beginning of editing cell
* Add data navigator
* Fill payroll with sample data
* Use SpreadsheetBindingManager component to bind data source properties to cells
* Assign binding source to DataSource properites of data navigator and spreadsheet binding manager
 
