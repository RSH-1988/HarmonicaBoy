using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace WindowsFormsApplication12
{
    public class ExportExcelFunctions
    {
        public int totalNumberOfRowsInExcel { get; set; }
        public  int fileExtension { get; set; }
        public Microsoft.Office.Interop.Excel.Workbook workbook { get; set; }
        public Microsoft.Office.Interop.Excel.Application app { get; set; }
        public Microsoft.Office.Interop.Excel.Worksheet w { get; set; }
        public Microsoft.Office.Interop.Excel.Range range { get; set; }
        public static String saveCsvFileIntoXlxFormat(String FilePath)
        {
            try
            {
                String NewPath = @"D:\AutRao\NewSheet1.xls";
                ExportExcelFunctions exp = new ExportExcelFunctions();
                exp.app = new Microsoft.Office.Interop.Excel.Application();
                object misValue = System.Reflection.Missing.Value;
                exp.workbook = exp.app.Workbooks.Open(FilePath, ReadOnly: true);
                exp.workbook.SaveAs(NewPath, XlFileFormat.xlUnicodeText, AccessMode: XlSaveAsAccessMode.xlNoChange);
                exp.workbook.Close();
                exp.app.Quit();
                //exp.w = exp.workbook.Worksheets.get_Item(1) as Microsoft.Office.Interop.Excel.Worksheet;
                // exp.w.SaveAs(@"D:\AutRao\NewSheet.xls", false, false);
                //exp.range = exp.w.UsedRange;
                return (NewPath);
            }
            catch (Exception ex) { throw ex; }
        }
        public static string openFileDialogBox()
        {
            
            string FilePath = null;
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();

                //Setting oroperties of dialog box
                openFileDialog1.Title = "Select file";
                openFileDialog1.InitialDirectory = @"D:\";
               // openFileDialog1.FileName = textBoxUpload.Text;
                openFileDialog1.Filter = "All Files(*.*)|*.*";
                openFileDialog1.FilterIndex = 1;
                openFileDialog1.RestoreDirectory = true;
                if (openFileDialog1.ShowDialog() == DialogResult.OK) // Test result.
                {
                   FilePath = openFileDialog1.FileName;
                    
                    
                }
                return FilePath;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static Microsoft.Office.Interop.Excel.Worksheet newExcel()
        {
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Application app;
            Microsoft.Office.Interop.Excel.Worksheet w;
            Microsoft.Office.Interop.Excel.Range range1;
            app = new Microsoft.Office.Interop.Excel.Application();
            if (app == null)
            {
                MessageBox.Show("Excel Could not started , Check that your office installation and project ref");
                return null;
            }
            app.Visible = true;
            workbook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            w = workbook.Worksheets.get_Item(1) as Microsoft.Office.Interop.Excel.Worksheet;
            range1 = w.UsedRange;
            return w;
        }
        public static Microsoft.Office.Interop.Excel.Worksheet openExcelInReadModeWithOutSheetName(String FilePath)
        {
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Application app;
            Microsoft.Office.Interop.Excel.Worksheet w;
            Microsoft.Office.Interop.Excel.Range range;
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                object misValue = System.Reflection.Missing.Value;
                workbook = app.Workbooks.Open(FilePath, ReadOnly: true);
                // Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true,
                //false, 0, true, false, false);
                w = workbook.Worksheets.get_Item(1) as Microsoft.Office.Interop.Excel.Worksheet;
                range = w.UsedRange;
                return (w);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please Close the file First! And Restart Uploading.");
                throw ex;
            }
        }
        public static System.Data.DataTable readHeader_InExcel(Microsoft.Office.Interop.Excel.Worksheet w,int rowNumber)
        {

            try
            {
                System.Data.DataTable dtHeader = new System.Data.DataTable();bool isEnd = false;
                DataRow dRowHeader;int i=1;
                dtHeader.Columns.Add("HeaderName", typeof(String));
                while (!isEnd)
                {
                    dRowHeader = dtHeader.NewRow();
                    string ss = (String)w.Cells[rowNumber, i].value2;
                    if (ss == null || ss.ToString().Length < 1) { isEnd = true; break; }
                    dRowHeader[0] = (String)w.Cells[rowNumber, i].value2; i++;
                    dtHeader.Rows.Add(dRowHeader);
                }
                return dtHeader;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static System.Data.DataTable readColumn_CheckDateTimeFormat(Microsoft.Office.Interop.Excel.Worksheet w,int rownumber,int columnnumber)
        {
            try
            {
                String ss = null;
                System.Data.DataTable dtColumn = new System.Data.DataTable();
                dtColumn.Columns.Add("dateFrom", typeof(DateTime)); dtColumn.Columns.Add("dateDueTo", typeof(DateTime));
                 CultureInfo ci = new CultureInfo("en-US");
                DataRow dtRow;
                int i=rownumber;bool isEnd = false;
                //*********
                while (!isEnd)
                {
                    ss = (String)w.Cells[i, 2].value2;
                    if (ss == null || ss.Length < 1)
                    {
                        isEnd = true; break;
                    }
                    dtRow = dtColumn.NewRow();
                    try
                    {
                        double flag;
                       
                        
                            dtRow[0] = DateTime.FromOADate(w.Cells[i, columnnumber].value2);
                        
                       
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            String ss1 = (String)w.Cells[i, columnnumber].value2;
                            if (ss1 != null)
                            {
                                DateTime dat = Convert.ToDateTime(ss1);
                                ss1 = dat.ToString(ci.NumberFormat);
                                dtRow[0] = Convert.ToDateTime(ss1);

                            }
                            else
                            {
                                dtRow[0] = DBNull.Value;

                            }
                        }
                        catch (Exception ex1)
                        {
                            String dateFrom = (String)w.Cells[i, columnnumber].value2;
                            double dateDouble = Convert.ToDouble(dateFrom);
                            dtRow[0] = DateTime.FromOADate(dateDouble);
                        }
                    }
                    try
                    {
                        double flag1;
                       
                            dtRow[1] = DateTime.FromOADate(w.Cells[i, columnnumber+1].value2);
                       
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            String ss1 = (String)w.Cells[i, columnnumber+1].value2;
                            //String ss2;
                            if (ss1.Contains('/') == true)
                            {
                                ss1 = ss1.Replace('/', '-');
                            }

                            if (ss1 != null)
                            {
                                DateTime dat = Convert.ToDateTime(ss1);
                                ss1 = dat.ToString(ci.NumberFormat);
                                dtRow[1] = Convert.ToDateTime(ss1);

                            }
                            else
                            {
                                dtRow[1] = DBNull.Value;

                            }
                        }
                        catch (Exception ex1)
                        {
                            String dateFrom = (String)w.Cells[i, columnnumber+1].value2;
                            double dateDouble = Convert.ToDouble(dateFrom);
                            dtRow[1] = DateTime.FromOADate(dateDouble);
                        }
                    }
                    dtColumn.Rows.Add(dtRow); i++;
                }
                return dtColumn;
            }
            catch(Exception ex)
            {
                throw(new Exception("Still not able to convert in dateformat",ex));
            }


        }
        public static Microsoft.Office.Interop.Excel.Worksheet openExcelInReadMode(String FilePath, String SheetName)
        {
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Application app;
            Microsoft.Office.Interop.Excel.Worksheet w;
            Microsoft.Office.Interop.Excel.Range range;
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                object misValue = System.Reflection.Missing.Value;
                workbook = app.Workbooks.Open(FilePath, ReadOnly: true);
                // Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true,
                //false, 0, true, false, false);
                w = workbook.Worksheets.get_Item(SheetName) as Microsoft.Office.Interop.Excel.Worksheet;
                range = w.UsedRange;
                return(w);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please Close the file First! And Restart Uploading.");
                throw ex;
            }
        }
        public static Microsoft.Office.Interop.Excel.Worksheet openExcelInWriteMode(String FilePath, String SheetName)
        {
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Application app;
            Microsoft.Office.Interop.Excel.Worksheet w;
            Microsoft.Office.Interop.Excel.Range range;
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                object misValue = System.Reflection.Missing.Value;
                workbook = app.Workbooks.Open(FilePath, ReadOnly: false);
                // Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true,
                //false, 0, true, false, false);
                w = workbook.Worksheets.get_Item(SheetName) as Microsoft.Office.Interop.Excel.Worksheet;
                range = w.UsedRange;
                return (w);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please Close the file First! And Restart Exporting.");
                throw ex;
            }
        }
        public static ExportExcelFunctions openExcelInWriteModeWithoutSHeetName(String FilePath)
        {
            
            try
            {
                ExportExcelFunctions exp = new ExportExcelFunctions();
                exp.app = new Microsoft.Office.Interop.Excel.Application();
                object misValue = System.Reflection.Missing.Value;
                exp.workbook = exp.app.Workbooks.Open(FilePath, ReadOnly: false);

                exp.w = exp.workbook.Worksheets.get_Item(1) as Microsoft.Office.Interop.Excel.Worksheet;
               // exp.w.SaveAs(@"D:\AutRao\NewSheet.xls", false, false);
                exp.range = exp.w.UsedRange;
                return (exp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please Close the file First! And Restart Exporting.");
                throw ex;
            }
        }
        public  string RangeAddress(Excel.Range rng)
        {
            bool missing = false;
            return rng.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1,
                   missing, missing);
        }
        public  string CellAddress(Excel.Worksheet sht, int row, int col)
        {
            return RangeAddress(sht.Cells[row, col]);
        }
        private void readTotalRows(Microsoft.Office.Interop.Excel.Worksheet w,int rowcounter)
        {
            int i = rowcounter;
            try
            {
                bool isEnd = false;
                while (!isEnd)
                {
                    String ss = null;
                    try
                    {
                        Double dd = w.Cells[i, 2].value2;
                        ss = dd.ToString();

                    }
                    catch (Exception)
                    {
                        try
                        {
                            ss = (String)w.Cells[i, 2].value2;
                        }
                        catch (Exception)
                        {
                            ss = w.Cells[i, 2].value2;
                        }
                    }
                    if (ss == null || ss.Length < 1)
                    {
                        isEnd = true; break;
                    }
                    i++;
                }
                totalNumberOfRowsInExcel = i;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static  int getIndexofNextServiceDateInSR(Microsoft.Office.Interop.Excel.Worksheet w, int rownumber, String columName)
        {
            try
            {
                int index = 0;
                int totalColumns = int.Parse(w.UsedRange.Columns.Count.ToString());
                for (int j = 1; j <= totalColumns; j++)
                {
                    String ss;
                    try
                    {
                        ss = (String)w.Cells[rownumber, j].value2;
                    }
                    catch (Exception)
                    {
                        ss = w.Cells[rownumber, j].value2;
                    }
                    if (ss != null)
                    {
                        if (columName.Trim().Equals(ss.ToString().Trim(), StringComparison.InvariantCultureIgnoreCase))
                        {
                            index = j;  break;
                        }
                    }
                }
                return index;
            }
            catch (Exception ex) { throw ex; }
        }
        public static int[] getIndexOfExcel(Microsoft.Office.Interop.Excel.Worksheet w, int rownumber, System.Data.DataTable ExcelFile)
        {
            try
            {

                int[] a = new int[100];
                for (int h = 0; h < 100; h++) { a[h] = 0; }
                
                int totalColumns = int.Parse(w.UsedRange.Columns.Count.ToString());
                for (int k = 0; k < ExcelFile.Columns.Count; k++)
                {
                    bool flag = false;
                    for (int j = 1; j <= totalColumns; j++)
                    {
                        String ss;
                        try
                        {
                            ss = (String)w.Cells[rownumber, j].value2;
                        }
                        catch (Exception)
                        {
                            ss = w.Cells[rownumber, j].value2;
                        }
                        if (ss != null)
                        {
                            if (ExcelFile.Columns[k].ColumnName.Trim().Equals(ss.ToString().Trim(),StringComparison.InvariantCultureIgnoreCase))
                            {
                                a[k] = j; flag = true; break;
                            }
                        }
                    }
                    if (flag == false)
                    {
                        a[k] = totalColumns + 2;
                    }
                }
                return a;
            }
            catch (Exception ex) { throw ex; }
        }
        public static int  insertNewColumnInExcel(Microsoft.Office.Interop.Excel.Worksheet w, int rowNumberOfColumnToBeShifted, int IndexOfColumnToBeShifted,String NewColumnName)
        {
            try
            {
                ExportExcelFunctions exp = new ExportExcelFunctions();
                String columnIndexToBeShifted = exp.CellAddress(w, rowNumberOfColumnToBeShifted, IndexOfColumnToBeShifted);
                Range r1; r1 = w.Range[columnIndexToBeShifted];
                Excel.Range col = r1.EntireColumn;
                col.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, false);
                w.Cells[1, IndexOfColumnToBeShifted] = NewColumnName.ToString();
                return IndexOfColumnToBeShifted;
            }
            catch (Exception ex) { throw ex; }
        }
        public static bool compareIntAndMonthValueOfDate(DateTime date, int monthValue)
        {
            try
            {
                if (int.Parse(date.Month.ToString()) == monthValue)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex) { throw ex; }
        }
        public static void applyFormulaForgettingMonth(Microsoft.Office.Interop.Excel.Worksheet w, int indexOfNewColumn, int indexOfColumnOnWhichFormulaApply, int rowdata,ExportExcelFunctions exp)
        {
            try
            {
                //ExportExcelFunctions exp = new ExportExcelFunctions();
                for (int j = rowdata; j <= exp.totalNumberOfRowsInExcel; j++)
                {
                    String columnNameOnWhichFormulaToBeApply = exp.CellAddress(w, j, indexOfColumnOnWhichFormulaApply);
                    String formulaForGettingMonth = "=CHOOSE(MONTH(" + columnNameOnWhichFormulaToBeApply.ToString() + "),\"1\",\"2\",\"3\",\"4\",\"5\",\"6\",\"7\",\"8\",\"9\",\"10\",\"11\",\"12\")";
                    w.Cells[j, indexOfNewColumn].formula = formulaForGettingMonth;
                }
            }
            catch (Exception ex) { throw ex; }
        }
        public static void applyFormulaForServiceInHondaServiceReminder(Microsoft.Office.Interop.Excel.Worksheet w,int rowData,String ColumnHeaderName,String NewColumnName,int rowHeader)
        {
            ExportExcelFunctions exp = new ExportExcelFunctions();
            exp.readTotalRows(w, rowData);
            int indexOfNextServiceDate;
            indexOfNextServiceDate = getIndexofNextServiceDateInSR(w, rowHeader, ColumnHeaderName);
           // indexOfTypeOfSerivce = getIndexofNextServiceDateInSR(w, 1, "Last Service Type");
           // indexOfDlrInvoiceDate = getIndexofNextServiceDateInSR(w, 1, "Dlr Invoice Date");
            //*******************************************************CodeForGettingMonthFromNextServiceDate**********************************************

            
            //*******************************************************NewCodeForGettingMonthFromDlrInvoiceDate***************************************************
            //indexOFNewColumn = insertNewColumnInExcel(w, 1, indexOfDlrInvoiceDate + 1,NewColumnName);
           //applyFormulaForgettingMonth(w, indexOFNewColumn, indexOfDlrInvoiceDate, rowData,exp);
           int indexOFNewColumn = insertNewColumnInExcel(w, rowHeader, indexOfNextServiceDate + 1, NewColumnName);
           applyFormulaForgettingMonth(w, indexOFNewColumn, indexOfNextServiceDate, rowData,exp);
           
            //*****************************************************End Here*******************************************************
          

        }
        public System.Data.DataTable universalExcelUpload(System.Data.DataTable ExcelFile, String FilePath, BackgroundWorker worker, int rowHeader, int rowData, int nonEmptyColumnIndex)
        {
            try
            {
                ExportExcelFunctions exp = new ExportExcelFunctions();
                String fileExtension = Path.GetExtension(FilePath);
                exp = openExcelInWriteModeWithoutSHeetName(FilePath);
                readTotalRows(exp.w, rowData);
             
                for (int k = 0; k < ExcelFile.Columns.Count; k++)
                {
                    if (ExcelFile.Columns[k].DataType.Name.Equals("DateTime"))
                    {
                        applyFormulaForServiceInHondaServiceReminder(exp.w, rowData, ExcelFile.Columns[k].ColumnName.ToString(), "Month Of "+ExcelFile.Columns[k].ColumnName.ToString(), rowHeader);
                    }
                }
                int[] index = getIndexOfExcel(exp.w, rowHeader, ExcelFile);
                CultureInfo ci = new CultureInfo("en-US");
                int i = rowData, count = 0; bool isEnd = false; DataRow dRow;
                while (!isEnd)
                {
                    String ss = null;
                    try
                    {
                        Double dd = exp.w.Cells[i, index[0]].value2;
                        ss = dd.ToString();

                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            ss = (String)exp.w.Cells[i, index[0]].value2;
                        }
                        catch (Exception)
                        {
                            try
                            {
                                int ii = exp.w.Cells[i, index[0]].value2;
                                ss = ii.ToString();
                            }
                            catch (Exception)
                            {
                                ss = exp.w.Cells[i, index[0]].value2;
                            }
                        }
                    }
                    if (ss == null || ss.ToString().Length <= 0)
                    {
                        isEnd = true; break;
                    }
                    dRow = ExcelFile.NewRow();
                    for (int j = 0; j < ExcelFile.Columns.Count; j++)
                    {
                        if (ExcelFile.Columns[j].DataType.Name.Equals("String") || ExcelFile.Columns[j].DataType.Name.Equals("string"))
                        {
                            try
                            {
                                dRow[j] = (String)exp.w.Cells[i, index[j]].value2;
                            }
                            catch (Exception)
                            {
                                try
                                {
                                    Double dd = exp.w.Cells[i, index[j]].value2;
                                    dRow[j] = dd.ToString();
                                }
                                catch (Exception)
                                {
                                    try
                                    {
                                        int ii = exp.w.Cells[i, index[j]].value2;
                                        dRow[j] = ii.ToString();
                                    }
                                    catch (Exception)
                                    {
                                        dRow[j] = exp.w.Cells[i, index[j]].value2;
                                    }
                                }
                            }
                        }
                        else if (ExcelFile.Columns[j].DataType.Name.Equals("int32") || ExcelFile.Columns[j].DataType.Name.Equals("Int32"))
                        {
                            try
                            {
                                dRow[j] = int.Parse(exp.w.Cells[i, index[j]].value2);
                            }
                            catch (Exception)
                            {
                                try
                                {
                                    dRow[j] = int.Parse((String)exp.w.Cells[i, index[j]].value2);
                                }
                                catch (Exception)
                                {
                                    try
                                    {
                                        dRow[j] = exp.w.Cells[i, index[j]].value2;
                                    }
                                    catch (Exception)
                                    {
                                        dRow[j] = DBNull.Value;
                                    }
                                }
                            }
                        }
                        else if (ExcelFile.Columns[j].DataType.Name.Equals("Decimal") || ExcelFile.Columns[j].DataType.Name.Equals("double"))
                        {
                            try
                            {
                                dRow[j] = decimal.Parse(exp.w.Cells[i, index[j]].value2);
                            }
                            catch (Exception)
                            {
                                try
                                {
                                    dRow[j] = decimal.Parse((String)exp.w.Cells[i, index[j]].value2);
                                }
                                catch (Exception)
                                {
                                    try
                                    {
                                        dRow[j] = exp.w.Cells[i, index[j]].value2;
                                    }
                                    catch (Exception)
                                    {
                                        dRow[j] = DBNull.Value;
                                    }
                                }
                            }
                        }
                        else if (ExcelFile.Columns[j].DataType.Name.Equals("DateTime"))
                        {
                            try
                            {

                                DateTime date = DateTime.FromOADate(exp.w.Cells[i, index[j]].value2);
                                string str;
                                int monthValueForDate = 0;                                                                 //getting Month value from next column to datetime column
                                monthValueForDate = int.Parse(exp.w.Cells[i, index[j] + 1].value2);
                               
                                    if (compareIntAndMonthValueOfDate(date, monthValueForDate))
                                    {
                                        str = date.ToString();
                                    }
                                    else
                                    {
                                        str = date.ToString("MM-dd-yyyy hh:MM:ss");
                                    }
                                dRow[j] = Convert.ToDateTime(str, null);
                              }
                            catch (Exception)
                            {
                                try
                                {
                                    String ss1 = (String)exp.w.Cells[i, index[j]].value2;
                                    DateTime newDate;
                                    if (ss1 == null || ss1.Equals("n/a") || ss1.Equals("."))
                                    {
                                        dRow[j] = DBNull.Value;
                                    }
                                    else if (ss1 != null)
                                    {
                                        if (ss1.Contains("\\"))
                                        {
                                            String[] dat1 = ss1.Split('\\');
                                          newDate = new DateTime(int.Parse(dat1[2].ToString()), int.Parse(dat1[1].ToString()), int.Parse(dat1[0].ToString()));
                                          ss1 = newDate.ToString();
                                        }
                                        else if (ss1.Contains("/"))
                                        {
                                            String[] dat1 = ss1.Split('/');
                                            newDate = new DateTime(int.Parse(dat1[2].ToString()), int.Parse(dat1[1].ToString()), int.Parse(dat1[0].ToString()));
                                            ss1 = newDate.ToString();
                                        }
                                        else if (ss1.Contains("-"))
                                        {
                                            String[] dat1 = ss1.Split('-');
                                            newDate = new DateTime(int.Parse(dat1[2].ToString()), int.Parse(dat1[1].ToString()), int.Parse(dat1[0].ToString()));
                                            ss1 = newDate.ToString();
                                        }
                                        DateTime dat = Convert.ToDateTime(ss1);
                                        ss1 = dat.ToString(ci.NumberFormat);
                                        dRow[j] = Convert.ToDateTime(ss1);
                                    }
                                    else
                                    {
                                        dRow[j] = DBNull.Value;
                                    }
                                }
                                catch (Exception ex1)
                                {
                                    throw ex1;
                                }
                            }
                        }
                    }//For loop ends here
                    if (string.IsNullOrEmpty(dRow.ItemArray[0].ToString()))
                    {
                        count++;
                        if (count >= 10)
                        {
                            isEnd = true;
                            break;
                        }
                    }
                    else
                    {
                        ExcelFile.Rows.Add(dRow);

                    }
                    // Thread.Sleep(100);
                    if (UtilityFunctions.PercentageCalculater(totalNumberOfRowsInExcel, i) < 100)
                        worker.ReportProgress(UtilityFunctions.PercentageCalculater(totalNumberOfRowsInExcel, i));
                    i++;
                }
                exp.workbook.Close(0); exp.app.Quit();
                //releaseObject(exp);
                if (ExcelFile.Columns.Contains("Dlr Invoice Date"))
                {
                    return UtilityFunctions.setServiceTypeInSRHiRise(ExcelFile);
                }
                else 
                if (ExcelFile.Columns.Contains("Job Card #"))
                {
                    return UtilityFunctions.addLastAndFirstNameinPSF(ExcelFile);
                }
                else
                return ExcelFile;
            }
            catch (Exception ex) { throw ex; }
        }
        public void exportCallTimeStatus(System.Data.DataTable CurrentStatus, DateTime From, DateTime To, String CallType)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Worksheet w = newExcel();
                if (w != null)
                {
                    Excel.Range rg, c1, c2;
                    if (CallType == null || CallType.ToString().Equals("Reminder"))
                    {
                        w.Cells[1, 1] = "Manual Reminder Call Time Status (" + From.ToString() + " - " + To.ToString() + ")";
                        c1 = w.Cells[1, 1];
                        c2 = w.Cells[1, 12];
                        rg = w.get_Range(c1, c2);
                        rg.Merge(true); rg.Interior.ColorIndex = 36; rg.Font.Bold = 20; rg.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    }
                    else if (CallType != null || CallType.ToString().Equals("Check Require"))
                    {
                        w.Cells[1, 1] = "Check Require Call Time Status (" + From.ToString() + " - " + To.ToString() + ")";
                        c1 = w.Cells[1, 1];
                        c2 = w.Cells[1, 12];
                        rg = w.get_Range(c1, c2);
                        rg.Merge(true); rg.Interior.ColorIndex = 36; rg.Font.Bold = 20; rg.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    }
                    else if (CallType != null || CallType.ToString().Equals("Check Bounce"))
                    {
                        w.Cells[1, 1] = "Check Bounce Call Time Status (" + From.ToString() + " - " + To.ToString() + ")";
                        c1 = w.Cells[1, 1];
                        c2 = w.Cells[1, 12];
                        rg = w.get_Range(c1, c2);
                        // rg = w.get_Range(w.Cells[1, 1], w.Cells[1, 11]);
                        rg.Merge(true); rg.Interior.ColorIndex = 36; rg.Font.Bold = 20; rg.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    }
                    int i = 2, j = 1;
                    for (int k = 1; k < CurrentStatus.Columns.Count; k++)
                    {

                        w.Cells[i, j++] = CurrentStatus.Columns[k].ColumnName.ToString();
                    }
                    c1 = w.Cells[2, 1];
                    c2 = w.Cells[2, j - 1];
                    rg = w.get_Range(c1, c2); rg.Font.Bold = 14; rg.Font.Color = Excel.XlThemeColor.xlThemeColorDark2; rg.Interior.ColorIndex = 3;
                    i = 3;
                    for (int k = 0; k < CurrentStatus.Rows.Count; k++, i++)
                    {
                        j = 1;
                        for (int p = 1; p < CurrentStatus.Columns.Count; p++, j++)
                        {
                            w.Cells[i, j] = CurrentStatus.Rows[k][p].ToString();
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
     
        public static System.Data.DataTable importSRHighRiseExcelSheet()
        {
            try
            {
                 Microsoft.Office.Interop.Excel.Worksheet w = openExcelInReadModeWithOutSheetName(openFileDialogBox());
                 System.Data.DataTable SRExcel = UtilityFunctions.createDataTableSRHighRise();
                CultureInfo ci = new CultureInfo("en-US");
                int i = 0;bool isEnd = false;
                DataRow dRow;
                while (!isEnd)
                {
                    string ss = (String)w.Cells[i + 2, 1].value2;
                    if (ss == null || ss.Length < 1)
                    {
                        isEnd = true; break;
                    }
                    dRow = SRExcel.NewRow();
                    try
                    {
                        dRow[0] = (String)w.Cells[i + 2, 1].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[0] = w.Cells[i + 2, 1].value2;
                    }
                    try
                    {
                        dRow[1] = (String)w.Cells[i + 2, 5].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[1] = w.Cells[i + 2, 5].value2; 
                    }
                    try
                    {
                        dRow[2] = (String)w.Cells[i + 2, 3].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[2] = w.Cells[i + 2, 3].value2;
                    }
                    try
                    {
                        
                        DateTime date = DateTime.FromOADate(w.Cells[i + 2, 16].value2);
                        string str = date.ToString();
                        dRow[3] = Convert.ToDateTime(str, null);

                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            String ss1 = (String)w.Cells[i + 2, 16].value2;
                            if (ss1 != null)
                            {
                                DateTime dat = Convert.ToDateTime(ss1);
                                ss1 = dat.ToString(ci.NumberFormat);
                                dRow[3] = Convert.ToDateTime(ss1);

                            }
                        }
                        catch (Exception ex1)
                        {
                            try
                            {
                                String ss2 = (String)w.Cells[i + 2, 63].value2;
                                ss2 = ss2.Replace("-", "/");
                                dRow[3] = DateTime.ParseExact(ss2.ToString(), "d/M/yyyy", ci);
                            }
                            catch (Exception ex2)
                            {
                                Double doubl = w.Cells[i + 2, 63].value2;
                               DateTime tDLR = DateTime.FromOADate(doubl);
                                String ss3 = tDLR.ToString(ci);
                                // ss3 = ss3.Replace("-", "/");
                                dRow[3] = Convert.ToDateTime(ss3.ToString(), ci);
                            }
                        }
                    }
                    try
                    {
                        dRow[4] = (String)w.Cells[i + 2, 19].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[4] = w.Cells[i + 2, 19].value2;
                    }
                    try
                    {
                        dRow[5] = (String)w.Cells[i + 2, 20].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[5] = w.Cells[i + 2, 20].value2;
                    }
                    String service = null;
                    try
                    {
                        service = (String)w.Cells[i + 2, 61].value2;
                    }
                    catch(Exception ex)
                    {
                        service = w.Cells[i + 2, 61].value2;
                        service = service.ToString();
                    }
                    if (service == null)
                    {
                        DateTime dtDLR;
                        try
                        {
                           dtDLR = DateTime.FromOADate(w.Cells[i + 2, 63]).value2;
                          // DateTime date = DateTime.FromOADate(w.Cells[i + 2, 19].value2);
                           string str = dtDLR.ToString();
                           dtDLR = Convert.ToDateTime(str, ci);
                        }
                        catch (Exception ex)
                        {

                            try
                            {
                                //String st = (String)w.Cells[i + 2, 63].value2;
                                //dtDLR = Convert.ToDateTime(st);
                                String ss1 = (String)w.Cells[i + 2, 63].value2;
                                DateTime dat = Convert.ToDateTime(ss1);
                                ss1 = dat.ToString(ci.NumberFormat);
                                dtDLR = Convert.ToDateTime(ss1, ci);
                            }
                            catch (Exception ex1)
                            {
                                try
                                {
                                    String ss2 = (String)w.Cells[i + 2, 63].value2;
                                    ss2 = ss2.Replace("-", "/");
                                    dtDLR = DateTime.ParseExact(ss2.ToString(), "d/M/yyyy", ci);
                                }
                                catch (Exception ex2)
                                {
                                    Double doubl = w.Cells[i + 2, 63].value2;
                                    dtDLR = DateTime.FromOADate(doubl);
                                    String ss3 = dtDLR.ToString(ci);
                                   // ss3 = ss3.Replace("-", "/");
                                    dtDLR = Convert.ToDateTime(ss3.ToString(), ci);
                                }

                            }
                         }
                        DateTime dtCurrent = DateTime.Today;
                        int days = int.Parse((dtCurrent - dtDLR).TotalDays.ToString());
                        if (days <= 31)
                            dRow[6] = "FREE 01";
                        else if (days <= 120)
                            dRow[6] = "FREE 02";
                        else if (days <= 240)
                            dRow[6] = "FREE 03";
                        else if (days <= 365)
                            dRow[6] = "FREE 04";
                        else if (days >= 365)
                            dRow[6] = "PAID";
                        else
                            dRow[6] = DBNull.Value;
                    }
                    else if (service.Equals("FREE 01"))
                    {
                        dRow[6] = "FREE 02";
                    }
                    else if (service.Equals("FREE 02"))
                    {
                        dRow[6] = "FREE 03";
                    }
                    else if (service.Equals("FREE 03"))
                    {
                        dRow[6] = "FREE 04";
                    }
                    else if (service.Equals("FREE 04"))
                    {
                        dRow[6] = "PAID";
                    }
                    else
                    {
                        dRow[6] = (String)w.Cells[i + 2, 61].value2;
                    }
                   
                    i++; SRExcel.Rows.Add(dRow);
                }
                return SRExcel;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static System.Data.DataTable importPSFHighRiseExcelSheet()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Worksheet w = openExcelInReadModeWithOutSheetName(openFileDialogBox());
                System.Data.DataTable PSFExcel = UtilityFunctions.createDataTablePSFHighRise();
                CultureInfo ci = new CultureInfo("en-US");
                int i = 0;bool isEnd = false;
                DataRow dRow;
                while (!isEnd)
                {
                    string ss = (String)w.Cells[i + 2, 1].value2;
                    if (ss == null || ss.Length < 1)
                    {
                        isEnd = true; break;
                    }
                    //String da = (String)w.Cells[i + 2, 18].value2;
                    //if (da == null || da.Length < 1)
                    //{
                    //    i++; continue;
                    //}
                    dRow = PSFExcel.NewRow();
                    dRow[0] = (String)w.Cells[i + 2, 7].value2 + " " + (String)w.Cells[i + 2, 6].value2;
                    try
                    {
                        dRow[1] = (String)w.Cells[i + 2, 20].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[1] = w.Cells[i + 2, 20].value2;
                    }
                    try
                    {
                        dRow[2] = (String)w.Cells[i + 2, 4].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[2] = w.Cells[i+2, 4].value2;
                    }
                    try
                    {
                        DateTime date = DateTime.FromOADate(w.Cells[i + 2, "SR Created Date/ Time"].value2);
                        string str = date.ToString();
                        dRow[4] = Convert.ToDateTime(str, ci);
                       
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            String ss1 = (String)w.Cells[i + 2, "SR Created Date/ Time"].value2;
                            if (ss1 != null)
                            {
                                DateTime dat = Convert.ToDateTime(ss1);
                                ss1 = dat.ToString(ci.NumberFormat);
                                dRow[4] = Convert.ToDateTime(ss1, ci);

                            }
                        }
                        catch (Exception ex1)
                        {
                            try
                            {
                                String ss2 = (String)w.Cells[i + 2, "SR Created Date/ Time"].value2;
                                ss2 = ss2.Replace("-", "/");
                                dRow[4] = DateTime.ParseExact(ss2.ToString(), "d/M/yyyy", ci);
                            }
                            catch (Exception ex2)
                            {
                                Double doubl = w.Cells[i + 2, "SR Created Date/ Time"].value2;
                                DateTime tDLR = DateTime.FromOADate(doubl);
                                String ss3 = tDLR.ToString(ci);
                                // ss3 = ss3.Replace("-", "/");
                                dRow[4] = Convert.ToDateTime(ss3.ToString(), ci);
                            }
                        }
                    }
                    try
                    {
                        dRow[3] = (String)w.Cells[i + 2, 3].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[3] = w.Cells[i + 2, 3].value2;
                    }
                    i++;
                    PSFExcel.Rows.Add(dRow);
                }
                return PSFExcel;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static System.Data.DataTable importJOBCARDCLOSEDSheet()
        {
            try
            {
                
                Microsoft.Office.Interop.Excel.Worksheet w = openExcelInReadModeWithOutSheetName(openFileDialogBox());
                System.Data.DataTable PSFExcel = UtilityFunctions.createDataTableJobCardPSFHighRise();
                CultureInfo ci = new CultureInfo("en-US");
                int i = 0;bool isEnd = false;
                DataRow dRow;
                while (!isEnd)
                {
                    string ss = (String)w.Cells[i + 2, 1].value2;
                    if (ss == null || ss.Length < 1)
                    {
                        isEnd = true; break;
                    }
                    dRow = PSFExcel.NewRow();
                    try
                    {
                        dRow[0] = (String)w.Cells[i + 2, 1].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[0] = w.Cells[i + 2, 1].value2;
                    }
                    try
                    {
                        dRow[1] = (String)w.Cells[i + 2, 2].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[1] = w.Cells[i + 2, 2].value2;
                    }
                    try
                    {
                        DateTime date = DateTime.FromOADate(w.Cells[i + 2, 8].value2);
                        string str = date.ToString();
                        dRow[2] = Convert.ToDateTime(str, ci);

                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            String ss1 = (String)w.Cells[i + 2, 8].value2;
                            if (ss1 != null)
                            {
                                DateTime dat = Convert.ToDateTime(ss1);
                                ss1 = dat.ToString(ci.NumberFormat);
                                dRow[2] = Convert.ToDateTime(ss1, ci);

                            }
                        }
                        catch (Exception ex1)
                        {
                            try
                            {
                                String ss2 = (String)w.Cells[i + 2, 8].value2;
                                ss2 = ss2.Replace("-", "/");
                                dRow[2] = DateTime.ParseExact(ss2.ToString(), "d/M/yyyy", ci);
                            }
                            catch (Exception ex2)
                            {
                                Double doubl = w.Cells[i + 2, 8].value2;
                                DateTime tDLR = DateTime.FromOADate(doubl);
                                String ss3 = tDLR.ToString(ci);
                                // ss3 = ss3.Replace("-", "/");
                                dRow[2] = Convert.ToDateTime(ss3.ToString(), ci);
                            }
                        }
                    }
                    try
                    {
                        dRow[3] = (String)w.Cells[i + 2, 9].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[3] = w.Cells[i + 2, 9].value2;
                    }
                    try
                    {
                        dRow[4] = (String)w.Cells[i + 2, 11].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[4] = w.Cells[i + 2, 11].value2;
                    }
                    try
                    {
                        dRow[5] = (String)w.Cells[i + 2, 12].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[5] = w.Cells[i + 2, 12].value2;
                    }
                    try
                    {
                        dRow[6] = (String)w.Cells[i + 2, 15].value2 +" "+ (String)w.Cells[i + 2, 14].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[6] = w.Cells[i + 2, 15].value2 +" "+w.Cells[i + 2, 14].value2;
                    }
                  
                    try
                    {
                        dRow[7] = (String)w.Cells[i + 2, 16].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[7] = w.Cells[i + 2, 16].value2;
                    }
                    try
                    {
                        dRow[8] = (String)w.Cells[i + 2, 28].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[8] = w.Cells[i + 2, 28].value2;
                    }
                    try
                    {
                        dRow[9] = (String)w.Cells[i + 2, 20].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[9] = w.Cells[i + 2, 20].value2;
                    }

                    i++; PSFExcel.Rows.Add(dRow);  
                }
                releaseObject(w);
                return PSFExcel;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static System.Data.DataTable importRandomExcel(Microsoft.Office.Interop.Excel.Worksheet w,System.Data.DataTable dt,System.Data.DataTable temp,System.Data.DataTable dtHeader, params int[] startPoint)
        {
            try
            {
                bool IsEnd = false;
                DataRow dRow = null; int i = 0;
                CultureInfo ci = new CultureInfo("en-US");
                for (i = startPoint[0]; !IsEnd ; i++)
                {
                    string ss=null;
                    try
                    {
                         ss = (String)w.Cells[i, 2].value2;
                    }
                    catch (Exception ex)
                    {
                        ss = w.Cells[i, 2].value2;
                    }
                    if (ss == null || ss.Length < 1)
                    {
                        IsEnd = true;
                        break;
                    }
                    dRow = dt.NewRow();
                    //Loop to read columns in Excel File
                    for (int j = 0,k=0; k < dtHeader.Rows.Count ; k++)
                    {
                        string fieldName = "";
                        if (j < temp.Rows.Count)
                        {
                            fieldName = temp.Rows[j][1].ToString();
                        }
                        bool flag=false;
                        if (dtHeader.Rows[k][0].ToString() == fieldName.ToString())     //Condition to check whether to skip or not in Excel File
                        {
                             flag = true;
                        }
                        else
                        {
                            continue;
                        }
                        if (flag)
                        {
                            string dataType = temp.Rows[j][2].ToString();
                            if (dataType.ToString() == "String")            //Read String type values
                            {
                                try
                                {
                                    dRow[j] = (String)w.Cells[i, k + 1].value2;
                                }
                                catch (Exception ex)
                                {
                                    dRow[j] = w.Cells[i, k + 1].value2;
                                }
                            }
                            else if (dataType.ToString() == "Decimal")          //Read Decimal Type values
                            {
                                try
                                {
                                    dRow[j] = decimal.Parse(w.Cells[i, k + 1].value2);
                                }
                                catch (Exception ex)
                                {
                                    dRow[j] = 0.0m;

                                }
                            }
                            else if (dataType.ToString() == "Datetime")         //Read Datetime type values
                            {
                                 
                                try
                                {
                                    DateTime date = DateTime.FromOADate(w.Cells[i, k + 1].value2);
                                    string str = date.ToString();
                                    dRow[j] = Convert.ToDateTime(str, null);
                                }
                                catch (Exception ex)
                                {
                                    try
                                    {
                                        String ss1 = (String)w.Cells[i, k + 1].value2;
                                        if (ss1 != null)
                                        {
                                            DateTime dat = Convert.ToDateTime(ss1);
                                            ss1 = dat.ToString(ci.NumberFormat);
                                            dRow[j] = Convert.ToDateTime(ss1);
                                        }
                                        else
                                        {
                                            dRow[j] = DBNull.Value;
                                        }
                                    }
                                    catch (Exception ex1)
                                    {
                                        try
                                        {
                                            String ss2 = (String)w.Cells[i, k + 1].value2;
                                            if (ss2 != null)
                                            {
                                                String ss2_t = ss2.Trim();
                                                DateTime da = DateTime.ParseExact(ss2_t, "d-M-yyyy", null);
                                                dRow[j] = da;
                                            }
                                            else
                                            {
                                                dRow[j] = DBNull.Value;
                                            }
                                        }
                                        catch (Exception ex2)
                                        {
                                            String ss3 = (String)w.Cells[i, k + 1].value2;
                                            if (ss3 != null)
                                            {
                                                String ss3_t = ss3.Trim();
                                                DateTime da1 = DateTime.ParseExact(ss3, "d-M-yyyy", CultureInfo.InvariantCulture);
                                                dRow[j] = da1;
                                            }
                                            else
                                            {
                                                dRow[j] = DBNull.Value;
                                            }
                                        }

                                    }

                                }

                            }
                            else if (dataType.ToString() == "Integer")      //Read Integer type values
                            {
                                try
                                {
                                    dRow[j] = int.Parse((String)w.Cells[i, k + 1].value2);
                                }
                                catch (Exception ex)
                                {
                                    dRow[j] = 0;

                                }
                            }
                            j++;
                        }
                    }
                    dt.Rows.Add(dRow); // Adding rows to Blank datatable
                }
                releaseObject(w);
                
                return dt; // Returning datatable
            }
            catch (Exception ex)
            {
                throw ex;
            }
          
        }
        public static System.Data.DataTable importHeroInsuranceSheetIntoHondaTable()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Worksheet w = openExcelInReadModeWithOutSheetName(openFileDialogBox());
                System.Data.DataTable ExcelFile = UtilityFunctions.createDataTableForHero();
                 CultureInfo ci = new CultureInfo("en-US");
                int i = 1;bool isEnd = false;
                DataRow dRow;
                while (!isEnd)
                {
                    string ss = (String)w.Cells[i + 1, 5].value2;
                    if (ss == null || ss.Length < 1)
                    {
                        isEnd = true; break;
                    }
                    dRow = ExcelFile.NewRow();
                    try
                    {
                        dRow[0] = (String)w.Cells[i + 1, 5].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[0] = w.Cells[i + 1, 5].value2;
                    }
                    try
                    {
                        dRow[1] = (String)w.Cells[i + 1, 21].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[1] = w.Cells[i + 1, 21].value2;
                    }
                    try
                    {
                        dRow[2] = (String)w.Cells[i + 1, 12].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[2] = w.Cells[i + 1, 12].value2;
                    }
                    try
                    {
                        DateTime date = DateTime.FromOADate(w.Cells[i + 1, 7].value2);
                        string str = date.ToString();
                        dRow[3] = Convert.ToDateTime(str, null);

                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            String ss1 = (String)w.Cells[i + 1, 7].value2;
                            if (ss1 != null)
                            {
                                DateTime dat = Convert.ToDateTime(ss1);
                                ss1 = dat.ToString(ci.NumberFormat);
                                dRow[3] = Convert.ToDateTime(ss1);

                            }
                        }
                        catch (Exception ex1)
                        {
                            try
                            {
                                String ss2 = (String)w.Cells[i + 1, 7].value2;
                                ss2 = ss2.Replace("-", "/");
                                dRow[3] = DateTime.ParseExact(ss2.ToString(), "d/M/yyyy", ci);
                            }
                            catch (Exception ex2)
                            {
                                Double doubl = w.Cells[i + 1, 7].value2;
                                DateTime tDLR = DateTime.FromOADate(doubl);
                                String ss3 = tDLR.ToString(ci);
                                // ss3 = ss3.Replace("-", "/");
                                dRow[3] = Convert.ToDateTime(ss3.ToString(), ci);
                            }
                        }
                    }
                    try
                    {
                        dRow[4] = (String)w.Cells[i + 1, 14].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[4] = w.Cells[i + 1, 14].value2;
                    }
                    try
                    {
                        dRow[5] = (String)w.Cells[i + 1, 10].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[5] = w.Cells[i + 1, 10].value2;
                    }
                    dRow[6] = DBNull.Value;
                    dRow[7] = DBNull.Value;
                    try
                    {
                        dRow[8] = (String)w.Cells[i + 1, 8].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[8] = w.Cells[i + 1, 8].value2;
                    }
                    try
                    {
                        dRow[9] = (String)w.Cells[i + 1, 9].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[9] = w.Cells[i + 1, 9].value2;
                    }
                    try
                    {
                        dRow[10] = (String)w.Cells[i + 1, 22].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[10] = w.Cells[i + 1, 22].value2;
                    }
                    dRow[11] = DBNull.Value;
                    ExcelFile.Rows.Add(dRow); i++;
                }
                releaseObject(w);
                return ExcelFile;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private static Object copyFromExcelToDataRowString(Object data)
        {
            Object copyHere = null;
            try
            {
                copyHere = (String)data;
            }
            catch (Exception ex)
            {
                copyHere = data;
            }
            return copyHere;
        }
        private static Object copyFromExcelToDataRowDecimal(Object data)
        {
            Object copyHere = null;
            try
            {
                copyHere = Decimal.Parse(data.ToString());
            }
            catch (Exception ex)
            {
                copyHere = data;
            }
            return copyHere;
        }
        private Object copyFromExcelToDataRowDateTime(Object data,CultureInfo ci)
        {
            Object copyHere = null;
            try
            {
                DateTime date = DateTime.FromOADate((Double)data);
                string str = date.ToString();
                copyHere = Convert.ToDateTime(str, null);

            }
            catch (Exception ex)
            {
                try
                {
                    String ss1 = data.ToString();
                    if (ss1 != null)
                    {
                        DateTime dat = Convert.ToDateTime(ss1);
                        ss1 = dat.ToString(ci.NumberFormat);
                        copyHere = Convert.ToDateTime(ss1);

                    }
                }
                catch (Exception ex1)
                {
                    try
                    {
                        String ss2 = data.ToString();
                        ss2 = ss2.Replace("-", "/");
                       copyHere = DateTime.ParseExact(ss2.ToString(), "d/M/yyyy", ci);
                    }
                    catch (Exception ex2)
                    {
                        Double doubl = Double.Parse(data.ToString());
                        DateTime tDLR = DateTime.FromOADate(doubl);
                        String ss3 = tDLR.ToString(ci);
                        // ss3 = ss3.Replace("-", "/");
                        copyHere = Convert.ToDateTime(ss3.ToString(), ci);
                    }
                }
            }
            return copyHere;

        }
        public static System.Data.DataTable importHeroPSFSheetIntoHondaTable()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Worksheet w = openExcelInReadModeWithOutSheetName(openFileDialogBox());
                System.Data.DataTable ExcelFile = UtilityFunctions.createDataTableForHeroPSFSheet();
                CultureInfo ci = new CultureInfo("en-US");
                int i = 0; bool isEnd = false;
                DataRow dRow;
                while (!isEnd)
                {
                    string ss = (String)w.Cells[i + 2, 2].value2;
                    if (ss == null || ss.Length < 1)
                    {
                        isEnd = true; break;
                    }
                    dRow = ExcelFile.NewRow();
                    dRow[0] = ((String)w.Cells[i + 2, 7].value2) + ((String)w.Cells[i + 2, 8].value2);
                    try
                    {
                        dRow[1] = (String)w.Cells[i + 2, 9].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[1] = w.Cells[i + 2, 9].value2;
                    }
                    try
                    {
                        dRow[2] = (String)w.Cells[i + 2, 2].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[2] = w.Cells[i + 2, 2].value2;
                    }
                    try
                    {
                        dRow[3] = (String)w.Cells[i + 2, 4].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[3] = w.Cells[i + 2, 4].value2;
                    }
                    try
                    {
                        dRow[4] = (String)w.Cells[i + 2, 5].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[4] = w.Cells[i + 2, 5].value2;
                    }
                    try
                    {
                        dRow[5] = (String)w.Cells[i + 2, 10].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[5] = w.Cells[i + 2, 10].value2;
                    }
                    try
                    {
                        DateTime date = DateTime.FromOADate(w.Cells[i + 2, 14].value2);
                        string str = date.ToString();
                        dRow[6] = Convert.ToDateTime(str, null);

                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            String ss1 = (String)w.Cells[i + 2, 14].value2;
                            if (ss1 != null)
                            {
                                DateTime dat = Convert.ToDateTime(ss1);
                                ss1 = dat.ToString(ci.NumberFormat);
                                dRow[6] = Convert.ToDateTime(ss1);

                            }
                        }
                        catch (Exception ex1)
                        {
                            try
                            {
                                String ss2 = (String)w.Cells[i + 2, 14].value2;
                                ss2 = ss2.Replace("-", "/");
                                dRow[6] = DateTime.ParseExact(ss2.ToString(), "d/M/yyyy", ci);
                            }
                            catch (Exception ex2)
                            {
                                Double doubl = w.Cells[i + 2, 14].value2;
                                DateTime tDLR = DateTime.FromOADate(doubl);
                                String ss3 = tDLR.ToString(ci);
                                // ss3 = ss3.Replace("-", "/");
                                dRow[6] = Convert.ToDateTime(ss3.ToString(), ci);
                            }
                        }
                    }
                    ExcelFile.Rows.Add(dRow); i++;
                }
                releaseObject(w);
                return ExcelFile;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static System.Data.DataTable importHeroServiceReminderSheetIntoHondaTable()
        {
            try
            {
                 Microsoft.Office.Interop.Excel.Worksheet w = openExcelInReadModeWithOutSheetName(openFileDialogBox());
                 System.Data.DataTable ExcelFile = UtilityFunctions.createDataTableForHeroReminderSheet();
                 CultureInfo ci = new CultureInfo("en-US");
                int i = 0;bool isEnd = false;
                DataRow dRow;
                while (!isEnd)
                {
                    string ss = (String)w.Cells[i + 3, 1].value2;
                    if (ss == null || ss.Length < 1)
                    {
                        isEnd = true; break;
                    }
                    dRow = ExcelFile.NewRow();
                    try
                    {
                        dRow[0] = (String)w.Cells[i + 3, 1].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[0] = w.Cells[i + 3, 1].value2;
                    }
                    dRow[1] = DateTime.Now;
                    try
                    {
                        dRow[2] = (String)w.Cells[i + 3, 6].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[2] = w.Cells[i + 3, 6].value2;
                    }
                    try
                    {
                       String Service = (String)w.Cells[i + 3, 14].value2;
                       if (Service.Equals("FSC"))
                       {
                           int Count;
                           try
                           {
                               Count = (int)w.Cells[i + 3, 52].value2;
                               
                           }
                           catch (Exception ex1)
                           {
                               try
                               {
                                   Count = w.Cells[i + 3, 52].value2;
                               }
                               catch (Exception ex2)
                               {
                                   Count = 7;
                               }
                              
                           }
                           switch (Count)
                           {
                               case 0:
                                   dRow[3] = "FREE 01";break;
                               case 1:
                                   dRow[3] = "FREE 02"; break;
                               case 2:
                                   dRow[3] = "FREE 03"; break;
                               case 3:
                                   dRow[3] = "FREE 04"; break;
                               case 4:
                                   dRow[3] = "FREE 05"; break;
                               case 5:
                                   dRow[3] = "FREE 06"; break;
                               case 6:
                                   dRow[3] = "PAID";break;
                               default:
                                   dRow[3] = "PAID";break;
                           }
                       }
                       else if (Service.Equals("Paid Service"))
                       {
                           dRow[3] = "PAID";
                       }
                       else
                       {
                           dRow[3] = Service.ToString();
                       }
                    }
                    catch (Exception ex)
                    {
                        dRow[3] = w.Cells[i + 3, 6].value2;
                    }
                    try
                    {
                        DateTime date = DateTime.FromOADate(w.Cells[i + 3, 15].value2);
                        string str = date.ToString();
                        dRow[4] = Convert.ToDateTime(str, null);

                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            String ss1 = (String)w.Cells[i + 3, 15].value2;
                            if (ss1 != null)
                            {
                                DateTime dat = Convert.ToDateTime(ss1);
                                ss1 = dat.ToString(ci.NumberFormat);
                                dRow[4] = Convert.ToDateTime(ss1);

                            }
                        }
                        catch (Exception ex1)
                        {
                            try
                            {
                                String ss2 = (String)w.Cells[i + 3, 15].value2;
                                ss2 = ss2.Replace("-", "/");
                                dRow[4] = DateTime.ParseExact(ss2.ToString(), "d/M/yyyy", ci);
                            }
                            catch (Exception ex2)
                            {
                                Double doubl = w.Cells[i + 3, 15].value2;
                                DateTime tDLR = DateTime.FromOADate(doubl);
                                String ss3 = tDLR.ToString(ci);
                                // ss3 = ss3.Replace("-", "/");
                                dRow[4] = Convert.ToDateTime(ss3.ToString(), ci);
                            }
                        }
                    }
                    try
                    {
                        dRow[5] = (String)w.Cells[i + 3, 17].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[5] = w.Cells[i + 3, 17].value2;
                    }
                    dRow[6] = ((String)w.Cells[i + 3, 27].value2) + ((String)w.Cells[i + 3, 28].value2);
                    try
                    {
                        dRow[7] = (String)w.Cells[i + 3, 29].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[7] = w.Cells[i + 3, 29].value2;
                    }
                    try
                    {
                        dRow[8] = (String)w.Cells[i + 3, 30].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[8] = w.Cells[i + 3, 30].value2;
                    }
                    try
                    {
                        dRow[9] = (String)w.Cells[i + 3, 51].value2;
                    }
                    catch (Exception ex)
                    {
                        dRow[9] = w.Cells[i + 3, 51].value2;
                    }
                    ExcelFile.Rows.Add(dRow); i++;
                }
                releaseObject(w);
                return ExcelFile;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
        }
        public static int exportToExcel(String FilePath,String SheetName,String SavingPath,System.Data.DataTable dt, int startRowIndex, int startColumnIndex,params int[] leavingColIndex)
        {
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Application app;
            Microsoft.Office.Interop.Excel.Worksheet w;
            Microsoft.Office.Interop.Excel.Range range;
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                object misValue = System.Reflection.Missing.Value;
                workbook = app.Workbooks.Open(FilePath, ReadOnly: false);
                // Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true,
                //false, 0, true, false, false);
                w = workbook.Worksheets.get_Item(SheetName) as Microsoft.Office.Interop.Excel.Worksheet;
                range = w.UsedRange;
               
                int StartColumnIndex = startColumnIndex;
               
                try
                {
                    for (int k = 0; k <= dt.Rows.Count - 1; k++)
                    {
                        startColumnIndex = StartColumnIndex;
                        for (int j = 0; j <= dt.Columns.Count - 1; j++)
                        {

                            if (j + startColumnIndex == leavingColIndex[0] || j + startColumnIndex == leavingColIndex[1])
                            {
                                if ((dt.Rows[k][21].ToString() == "YES" || dt.Rows[k][21].ToString() == "Yes" || dt.Rows[k][21].ToString() == "yes") && j + startColumnIndex == leavingColIndex[0])
                                    w.Cells[k + startRowIndex, j + startColumnIndex] = 4;
                                else if ((dt.Rows[k][21].ToString() == "NO" || dt.Rows[k][21].ToString() == "No" || dt.Rows[k][21].ToString() == "no") && j + startColumnIndex == leavingColIndex[0])
                                    w.Cells[k + startRowIndex, j + startColumnIndex] = 1;
                                else {  }
                                 startColumnIndex += leavingColIndex.Length;
                                 String data = dt.Rows[k][j].ToString(); 
                                 w.Cells[k + startRowIndex, j + startColumnIndex] = data;
                                //continue;
                            }
                            else if (j + startColumnIndex == 21)
                            {
                                if ((dt.Rows[k][18] != null && dt.Rows[k][18].ToString().Length > 0))
                                    //&& (dt.Rows[k][21] != null &&
                                    //dt.Rows[k][21].ToString() == "YES" || dt.Rows[k][21].ToString() == "Yes" || dt.Rows[k][21].ToString() == "yes"))
                                {
                                    w.Cells[k + startRowIndex, j + startColumnIndex] = "YES"; 
                                }
                                else 
                                {
                                    w.Cells[k + startRowIndex, j + startColumnIndex] = "NO";
                                }
                                startColumnIndex += 1;
                                String data = dt.Rows[k][j].ToString();
                                w.Cells[k + startRowIndex, j + startColumnIndex] = data;
                            }
                            else if (j + startColumnIndex == 27)
                            {
                                if ((dt.Rows[k][18] != null && dt.Rows[k][18].ToString().Length > 0)&& (dt.Rows[k][21] != null &&
                                dt.Rows[k][21].ToString() == "YES" || dt.Rows[k][21].ToString() == "Yes" || dt.Rows[k][21].ToString() == "yes"))
                                {
                                    w.Cells[k + startRowIndex, j + startColumnIndex] = "YES";
                                }
                                else
                                {
                                    w.Cells[k + startRowIndex, j + startColumnIndex] = "NO";
                                }
                                //startColumnIndex += 1;
                                //String data = dt.Rows[k][j].ToString();
                                //w.Cells[k + startRowIndex, j + startColumnIndex] = data;
                            }
                            else
                            {
                                String data = dt.Rows[k][j].ToString();
                                w.Cells[k + startRowIndex, j + startColumnIndex] = data;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                workbook.SaveAs(SavingPath);
                //, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                workbook.Close(true, misValue, misValue);
                app.Quit();

                releaseObject(w);
                releaseObject(workbook);
                releaseObject(app);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return 1;
        }
        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        /* for (int i = 2; i <= exp.totalNumberOfRowsInExcel; i++)
         {
             String ss = (String)w.Cells[i, indexOfTypeOfSerivce].value2;
             if (ss == null || ss.ToString().Equals("") || ss.ToString().Length <= 0)
             {
                // String columnIndexForService = exp.CellAddress(w, i, indexOfTypeOfSerivce);
                 String coulmnIndexForDate = exp.CellAddress(w, i, indexOfDlrInvoiceDate);
                 String formulaSecond = "=IF((TODAY()-" + coulmnIndexForDate.ToString() + ")<=30,\"FREE 01\",IF((TODAY()-" + coulmnIndexForDate.ToString() + ")<=120,\"FREE 02\",IF((TODAY()-" + coulmnIndexForDate.ToString() + ")<=240,\"FREE 03\",IF((TODAY()-" + coulmnIndexForDate.ToString() + ")<=356,\"FREE 04\",\"PAID\"))))";
                 w.Cells[i, indexOfTypeOfSerivce].formula = formulaSecond;
                   
             }
             else if (ss.ToString().Equals("FREE 01"))
             {
                 w.Cells[i, indexOfTypeOfSerivce] = "FREE 02";
             }
             else if (ss.ToString().Equals("FREE 02"))
             {
                 w.Cells[i, indexOfTypeOfSerivce] = "FREE 03";
             }
             else if (ss.ToString().Equals("FREE 03"))
             {
                 w.Cells[i, indexOfTypeOfSerivce] = "FREE 04";
             }
             else if (ss.ToString().Equals("FREE 04"))
             {
                 w.Cells[i, indexOfTypeOfSerivce] = "PAID";
             }
             else { }
                
         }*/

        //}

        // string str;
        //if (indexOfNextServiceDate != 0 && index[j].Equals(indexOfNextServiceDate))
        //{
        //    if (date.Month.Equals(monthValue))
        //    {
        //        str = date.ToString();

        //    }
        //    else
        //    {
        //        DateTime newDate = new DateTime(date.Year, date.Day, date.Month);
        //        str = newDate.ToString();
        //    }
        //}
        //else
        //{
    }
}
