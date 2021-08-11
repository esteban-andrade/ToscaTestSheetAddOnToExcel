using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Tricentis.TCAddOns;
using Tricentis.TCAPIObjects.Objects;
using DataTable = System.Data.DataTable;

namespace TestSheetAddOn
{
    public class TestSheetAddOn : TCAddOn
    {
        public override string UniqueName
        {
            get { return "Testsheet Helper"; }
        }
    }
    public class ExportTCtoExcel : TCAddOnTask
    {
        public override string Name => "Export Test Design model to Excel";

        public override Type ApplicableType => typeof(TestSheet);

        //Check if it's a testcase folder not just any folder
        public override bool IsTaskPossible(TCObject obj) => true;
        public override bool RequiresChangeRights => true;

        public override TCObject Execute(TCObject objectToExecuteOn, TCAddOnTaskContext taskContext)
        {
            if (objectToExecuteOn is TestSheet)
            {
                //Excel information            
                int headerrow = 1;


                Microsoft.Office.Interop.Excel.Application oXL;
                try
                {
                    oXL = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    oXL.Visible = false;
                }
                catch
                {
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oXL.Visible = false;
                }


                TestSheet ts = (TestSheet)objectToExecuteOn;
                string shname = ts.DisplayedName;
                string[] headervalues = { };
                int intcounter = 0;

                //Get Excel objects
                Workbook oWB = oXL.Workbooks.Add();
                Worksheet oWS = oWB.Sheets.Add();
                //oWS.Name = shname;

                //Find all the instances
                IEnumerable<TDElementWrapper> elementWrappers = TestsheetElementsHelper.GetElementsRecursively(ts);

                var instances = ts.Instances.Items.ToArray();

                int renamecounter = 1;
                int rowcountt = 0;

                //Write header rows
                int colcounter = 2;
                oWS.Cells[headerrow, 1] = "Instances";
                foreach (TDInstance tdheader in ts.Instances.Items)
                {

                    for (int i = 0; i < elementWrappers.Count(); i++)
                    {
                        TDElementWrapper ewvarheader = elementWrappers.ElementAt(i);
                        if (ewvarheader != null)
                        {
                            string path = ewvarheader.Path;
                            if (path.Contains('.'))
                            {
                                string[] patharray = path.Split('.');
                                for (int j = 0; j < patharray.Length; j++)
                                {
                                    oWS.Cells[headerrow + (j * 1), colcounter] = patharray[j];
                                    if (rowcountt < headerrow + (j * 1)) { rowcountt = headerrow + (j * 1); }
                                }
                            }
                            else
                            {
                                oWS.Cells[headerrow, colcounter] = path;
                            }
                            colcounter = colcounter + 1;
                        }
                    }
                    break;
                }

                Range oHeader = oWS.Range[oWS.Cells[headerrow, 1], oWS.Cells[rowcountt, colcounter - 1]];
                int contentrow = rowcountt + 1;

                //Write rows -- Instances names
                foreach (TDInstance tdi in ts.Instances.Items)
                {
                    oWS.Cells[contentrow, 1] = tdi.DisplayedName;
                    colcounter = 2;

                    for (int i = 0; i < elementWrappers.Count(); i++)
                    {
                        TDElementWrapper ewvar = elementWrappers.ElementAt(i);
                        if (ewvar != null)
                        {
                            //tdi.Name += ElementToValueMapper.GetInstanceValueStringForElement(tdi, ew);
                            string path = ewvar.Path;
                            oWS.Cells[contentrow, colcounter] = ElementToValueMapper.GetInstanceValueStringForElement(tdi, ewvar);
                            colcounter = colcounter + 1;
                        }
                    }

                    contentrow = contentrow + 1;
                }

                //Try to get the relationships


                //Format excel document
                oWS.Columns.AutoFit();
                oWS.Columns.Font.Size = 11;
                oWS.Columns.Font.Name = "Calibri";
                Range myCell = oWS.Range[oWS.Cells[rowcountt + 1, 2], oWS.Cells[rowcountt + 1, 2]];
                myCell.Activate();
                myCell.Application.ActiveWindow.FreezePanes = true;
                oHeader.Columns.Font.Color = XlRgbColor.rgbWhite;
                oHeader.Columns.Font.Bold = true;
                oHeader.Columns.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                oHeader.Columns.Interior.Color = XlRgbColor.rgbDarkSlateGrey;
                Range contents = oWS.Range[oWS.Cells[rowcountt + 1, 1], oWS.Cells[contentrow - 1, colcounter - 1]];
                contents.WrapText = true;
                contents.Borders.LineStyle = XlLineStyle.xlContinuous;
                contents.Borders.Weight = XlBorderWeight.xlThin;
                contents.Borders.ColorIndex = XlColorIndex.xlColorIndexAutomatic;

                //Save document
                System.IO.Directory.CreateDirectory("C:\\Tosca_Projects\\Tosca_Exports");
                oWB.SaveAs("C:\\Tosca_Projects\\Tosca_Exports\\" + shname + "_TCD Extract " + DateTime.Now.ToString("MMddyyyy_hhmmss"), Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //Close Excel application
                oWB.Close(true);
                oXL.Quit();

                //Release Excel objects
                Marshal.ReleaseComObject(oWS);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);

                Process.Start(@"C:\Tosca_Projects\Tosca_Exports\");
            }
            return null;
        }

    }

    public class ExportTCtoExcelForJira : TCAddOnTask
    {
        public override string Name => "Export TCD to Excel for Jira";

        public override Type ApplicableType => typeof(TestSheet);

        //Check if it's a testcase folder not just any folder
        //public override bool IsTaskPossible(TCObject obj) => true;
        //public override bool RequiresChangeRights => true;

        public override TCObject Execute(TCObject objectToExecuteOn, TCAddOnTaskContext taskContext)
        {
            if (objectToExecuteOn is TestSheet)
            {
                //Excel information            
                int headerrow = 2;
                

                Microsoft.Office.Interop.Excel.Application oXL;
                try
                {
                    oXL = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    oXL.Visible = false;
                }
                catch
                {
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oXL.Visible = false;
                }
               

                TestSheet ts = (TestSheet)objectToExecuteOn;
                string shname = ts.DisplayedName;
                string[] headervalues = { };
                int intcounter = 0;

                //Get Excel objects
                Workbook oWB = oXL.Workbooks.Add();
                Worksheet oWS = oWB.Sheets.Add();
                //oWS.Name = shname;
                
                //Find all the instances
                IEnumerable<TDElementWrapper> elementWrappers = TestsheetElementsHelper.GetElementsRecursively(ts);
                                
                var instances = ts.Instances.Items.ToArray();

                int renamecounter = 1;
                int rowcountt = 0;

                //Write header rows
                int colcounter = 2;
                int headerrow_names = 1;
                oWS.Cells[headerrow_names, 1] = "Key";
                oWS.Cells[headerrow_names, 2] = "Name";
                oWS.Cells[headerrow_names, 3] = "Status";
                oWS.Cells[headerrow_names, 4] = "Precondition";
                oWS.Cells[headerrow_names, 5] = "Objective";
                oWS.Cells[headerrow_names, 6] = "Folder";
                oWS.Cells[headerrow_names, 7] = "Priority";
                oWS.Cells[headerrow_names, 8] = "Component";
                oWS.Cells[headerrow_names, 9] = "Labels";
                oWS.Cells[headerrow_names, 10] = "Owner";
                oWS.Cells[headerrow_names, 11] = "Estimated Time";
                oWS.Cells[headerrow_names, 12] = "Coverage (Issues)";
                oWS.Cells[headerrow_names, 13] = "Coverage (Pages)";
                oWS.Cells[headerrow_names, 14] = "Test Script (Step-by-Step) - Step";
                oWS.Cells[headerrow_names, 15] = "Test Script (Step-by-Step) - Test Data";
                oWS.Cells[headerrow_names, 16] = "Test Script (Step-by-Step) - Expected Result";
                oWS.Cells[headerrow_names, 17] = "Test Script (Plain Text))";
                oWS.Cells[headerrow_names, 18] = "Test Script (BDD)";
                
                int contentrow = headerrow;

                //Write rows -- Instances names
                foreach (TDInstance tdi in ts.Instances.Items)
                {
                    string displayName = ElementToValueMapper.GetInstanceValueStringForElement(tdi, elementWrappers.ElementAt(0));
                    oWS.Cells[contentrow, 2] = tdi.DisplayedName;
                    oWS.Cells[contentrow, 3] = "Approved";
                    oWS.Cells[contentrow, 7] = "High";
                    oWS.Cells[contentrow, 8] = "PCC";
                    oWS.Cells[contentrow, 9] = "Drop 1 Collection Strategies";
                    //oWS.Cells[contentrow, 14] = displayName;

                    //colcounter = 2;
                    string precond = "";
                    string process = "";
                    string verification = "";
                    for (int i = 0; i < elementWrappers.Count(); i++)
                    {
                        TDElementWrapper ewvar = elementWrappers.ElementAt(i);
                        if (ewvar != null)
                        {
                            string path = ewvar.Path;
                            //Concatenates all the PreCondition and Process instance values in a string
                            if(path.Contains("Precondition") )
                            {
                                string value = ElementToValueMapper.GetInstanceValueStringForElement(tdi, ewvar);
                                if(value != null)
                                {
                                    string[] patharray = path.Split('.');       
                                    if(precond == "")
                                    {
                                        precond = patharray[patharray.Length - 1] + ": " + value;
                                    }
                                    else
                                    {
                                        precond = precond + "\n" + patharray[patharray.Length - 1] + ": " + value;
                                    }
                                    
                                }
                            }

                            else if (path.Contains("Process"))
                            {
                                string value = ElementToValueMapper.GetInstanceValueStringForElement(tdi, ewvar);
                                if (value != null)
                                {
                                    string[] patharray = path.Split('.');
                                    if (precond == "")
                                    {
                                        process = patharray[patharray.Length - 1] + ": " + value;
                                    }
                                    else
                                    {
                                        process = process + "\n" + patharray[patharray.Length - 1] + ": " + value;
                                    }

                                }

                            }
                            //Concatenates all the Verification instance values in a string
                            else if (path.Contains("Verification"))
                            {
                                string value = ElementToValueMapper.GetInstanceValueStringForElement(tdi, ewvar);
                                if (value != null)
                                {
                                    string[] patharray = path.Split('.');
                                    if(verification == "")
                                    {
                                        verification = patharray[patharray.Length - 1] + ": " + value;
                                    }
                                    else
                                    {
                                        verification = verification + "\n" + patharray[patharray.Length - 1] + ": " + value;
                                    }
                                }
                            } 
                        }
                    }
                    oWS.Cells[contentrow, 4] = displayName + "\n" + precond;
                    oWS.Cells[contentrow, 16] = displayName + "\n" + verification;
                    oWS.Cells[contentrow, 14] = displayName + "\n" + process;

                    contentrow = contentrow + 1;
                }

                //Try to get the relationships
                

                //Format excel document
                oWS.Columns.AutoFit();
                oWS.Columns.Font.Size = 11;
                oWS.Columns.Font.Name = "Calibri";

                //Save document
                System.IO.Directory.CreateDirectory("C:\\Tosca_Projects\\Tosca_Exports");
                oWB.SaveAs("C:\\Tosca_Projects\\Tosca_Exports\\"+ shname + "_TCD Extract for Jira " + DateTime.Now.ToString("MMddyyyy_hhmmss"), Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //Close Excel application
                oWB.Close(true);
                oXL.Quit();

                //Release Excel objects
                Marshal.ReleaseComObject(oWS);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);

                Process.Start(@"C:\Tosca_Projects\Tosca_Exports\");
            }
            return null;
        }
        
    }

    public class ExportAttributetoExcel : TCAddOnTask
    {
        public override string Name => "Export TD Attribute model to Excel";

        public override Type ApplicableType => typeof(TDAttribute);

        //Check if it's a testcase folder not just any folder
        public override bool IsTaskPossible(TCObject obj) => true;
        public override bool RequiresChangeRights => true;

        public override TCObject Execute(TCObject objectToExecuteOn, TCAddOnTaskContext taskContext)
        {
            if (objectToExecuteOn is TDAttribute)
            {
                //Excel information            
                int headerrow = 1;


                Microsoft.Office.Interop.Excel.Application oXL;
                try
                {
                    oXL = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    oXL.Visible = false;
                }
                catch
                {
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oXL.Visible = false;
                }


                TDAttribute ts = (TDAttribute)objectToExecuteOn;
                string shname = ts.DisplayedName;
                string[] headervalues = { };
                int intcounter = 0;

                //Get Excel objects
                Workbook oWB = oXL.Workbooks.Add();
                Worksheet oWS = oWB.Sheets.Add();
                //oWS.Name = shname;

                //Find all the instances
                IEnumerable<TDElementWrapper> elementWrappers = TestsheetElementsHelper.GetAttrRecursively(ts);

                var instances = ts.Instances.Items.ToArray();

                int renamecounter = 1;
                int rowcountt = 0;

                //Write header rows
                int colcounter = 2;
                oWS.Cells[headerrow, 1] = "Instances";
                foreach (TDInstance tdheader in ts.Instances.Items)
                {

                    for (int i = 0; i < elementWrappers.Count(); i++)
                    {
                        TDElementWrapper ewvarheader = elementWrappers.ElementAt(i);
                        if (ewvarheader != null)
                        {
                            string path = ewvarheader.Path;
                            if (path.Contains('.'))
                            {
                                string[] patharray = path.Split('.');
                                for (int j = 0; j < patharray.Length; j++)
                                {
                                    oWS.Cells[headerrow + (j * 1), colcounter] = patharray[j];
                                    if (rowcountt < headerrow + (j * 1)) { rowcountt = headerrow + (j * 1); }
                                }
                            }
                            else
                            {
                                oWS.Cells[headerrow, colcounter] = path;
                            }
                            colcounter = colcounter + 1;
                        }
                    }
                    break;
                }

                Range oHeader = oWS.Range[oWS.Cells[headerrow, 1], oWS.Cells[rowcountt, colcounter - 1]];
                int contentrow = rowcountt + 1;

                //Write rows -- Instances names
                foreach (TDInstance tdi in ts.Instances.Items)
                {
                    oWS.Cells[contentrow, 1] = tdi.DisplayedName;
                    colcounter = 2;

                    for (int i = 0; i < elementWrappers.Count(); i++)
                    {
                        TDElementWrapper ewvar = elementWrappers.ElementAt(i);
                        if (ewvar != null)
                        {
                            //tdi.Name += ElementToValueMapper.GetInstanceValueStringForElement(tdi, ew);
                            string path = ewvar.Path;
                            oWS.Cells[contentrow, colcounter] = ElementToValueMapper.GetInstanceValueStringForElement(tdi, ewvar);
                            colcounter = colcounter + 1;
                        }
                    }

                    contentrow = contentrow + 1;
                }

                //Try to get the relationships


                //Format excel document
                oWS.Columns.AutoFit();
                oWS.Columns.Font.Size = 11;
                oWS.Columns.Font.Name = "Calibri";
                Range myCell = oWS.Range[oWS.Cells[rowcountt + 1, 2], oWS.Cells[rowcountt + 1, 2]];
                myCell.Activate();
                myCell.Application.ActiveWindow.FreezePanes = true;
                oHeader.Columns.Font.Color = XlRgbColor.rgbWhite;
                oHeader.Columns.Font.Bold = true;
                oHeader.Columns.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                oHeader.Columns.Interior.Color = XlRgbColor.rgbDarkSlateGrey;
                Range contents = oWS.Range[oWS.Cells[rowcountt + 1, 1], oWS.Cells[contentrow - 1, colcounter - 1]];
                contents.WrapText = true;
                contents.Borders.LineStyle = XlLineStyle.xlContinuous;
                contents.Borders.Weight = XlBorderWeight.xlThin;
                contents.Borders.ColorIndex = XlColorIndex.xlColorIndexAutomatic;

                //Save document
                System.IO.Directory.CreateDirectory("C:\\Tosca_Projects\\Tosca_Exports");
                oWB.SaveAs("C:\\Tosca_Projects\\Tosca_Exports\\" + shname + "_TCD Extract " + DateTime.Now.ToString("MMddyyyy_hhmmss"), Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //Close Excel application
                oWB.Close(true);
                oXL.Quit();

                //Release Excel objects
                Marshal.ReleaseComObject(oWS);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);

                Process.Start(@"C:\Tosca_Projects\Tosca_Exports\");
            }
            return null;
        }

    }
    public class RenameInstancesAddOn : TCAddOnTask
    {
        public override Type ApplicableType
        {
            get { return typeof(TestSheet); }
        }

        public override TCObject Execute(TCObject objectToExecuteOn, TCAddOnTaskContext taskContext)
        {
            frmTDIName f = new frmTDIName();
            if (f.ShowDialog() == DialogResult.OK)
            {
                if (objectToExecuteOn is TestSheet)
                {
                    TestSheet ts = (TestSheet)objectToExecuteOn;

                    List<TDInstanceNameComponent> components = TestsheetElementsHelper.Initialise(f.NameFormat);
                    IEnumerable<TDElementWrapper> elementWrappers = TestsheetElementsHelper.GetElementsRecursively(ts);

                    var instances = ts.Instances.Items.ToArray();
                    int renamecounter = 1;
                    int errorCount = 0;
                    foreach (TDInstance tdi in ts.Instances.Items)
                    {
                        String oldName = tdi.Name;
                        String newName = "";

                        try
                        {
                            tdi.Name = ""; // Clear old name
                            foreach (TDInstanceNameComponent comp in components)
                            {
                                if (comp.IsPath)
                                {
                                    TDElementWrapper ew =
                                        elementWrappers.FirstOrDefault(
                                            e => e.Path.Equals(comp.Text, StringComparison.CurrentCultureIgnoreCase));
                                    if (ew != null)
                                    {
                                        //tdi.Name += ElementToValueMapper.GetInstanceValueStringForElement(tdi, ew);
                                        newName += ElementToValueMapper.GetInstanceValueStringForElement(tdi, ew);
                                    }
                                }
                                else
                                {
                                    //tdi.Name += comp.Text;
                                    newName += comp.Text;
                                }
                            }
                            //Loop to make sure that there are not continuous _                            
                            while (newName.Contains("__"))
                            {
                                newName = newName.Replace("__", "_");
                            }

                            //Check for the trailing character and omit it
                            if (newName.Substring(newName.Length - 1, 1) == "_")
                            {
                                newName = newName.Substring(0, newName.Length - 1);
                            }

                            //Check for name already exists
                            string newNamewithcount = "";
                            if (instances.Any(instance => instance.Name == newName))
                            {
                                newNamewithcount = newName + "_" + renamecounter.ToString("D3");
                                while (instances.Any(instance => instance.Name == newNamewithcount))
                                {

                                    renamecounter++;
                                    newNamewithcount = newName + "_" + renamecounter.ToString("D3");
                                }
                            }
                            else
                            {
                                newNamewithcount = newName;
                            }


                            tdi.Name = newNamewithcount;
                            //recurssive function to check name
                            //string finalname = RecursiveSearch(ref newName, ref ts);
                        }
                        catch (Exception)
                        {
                            // Reset to old name
                            tdi.Name = oldName;
                            errorCount++;
                        }
                    }

                    if (errorCount > 0)
                    {
                        MessageBox.Show(errorCount.ToString() + " instances could not be renamed.");
                    }
                }
            }

            return null;
        }


        public override string Name
        {
            get { return "Rename Instances"; }
        }

    }

    public class RenameAttributesAddOn : TCAddOnTask
    {
        public override Type ApplicableType
        {
            get { return typeof(TDAttribute); }
        }

        public override TCObject Execute(TCObject objectToExecuteOn, TCAddOnTaskContext taskContext)
        {
            frmTDIName f = new frmTDIName();
            if (f.ShowDialog() == DialogResult.OK)
            {
                if (objectToExecuteOn is TDAttribute)
                {
                    TDAttribute ts = (TDAttribute)objectToExecuteOn;

                    List<TDInstanceNameComponent> components = TestsheetElementsHelper.Initialise(f.NameFormat);
                    IEnumerable<TDElementWrapper> elementWrappers = TestsheetElementsHelper.GetAttrRecursively(ts);

                    var instances = ts.Instances.Items.ToArray();
                    int renamecounter = 1;
                    int errorCount = 0;
                    foreach (TDInstance tdi in ts.Instances.Items)
                    {
                        String oldName = tdi.Name;
                        String newName = "";

                        try
                        {
                            tdi.Name = ""; // Clear old name
                            foreach (TDInstanceNameComponent comp in components)
                            {
                                if (comp.IsPath)
                                {
                                    TDElementWrapper ew =
                                        elementWrappers.FirstOrDefault(
                                            e => e.Path.Equals(comp.Text, StringComparison.CurrentCultureIgnoreCase));
                                    if (ew != null)
                                    {
                                        //tdi.Name += ElementToValueMapper.GetInstanceValueStringForElement(tdi, ew);
                                        newName += ElementToValueMapper.GetInstanceValueStringForElement(tdi, ew);
                                    }
                                }
                                else
                                {
                                    //tdi.Name += comp.Text;
                                    newName += comp.Text;
                                }
                            }
                            //Loop to make sure that there are not continuous _                            
                            while (newName.Contains("__"))
                            {
                                newName = newName.Replace("__", "_");
                            }

                            //Check for the trailing character and omit it
                            if (newName.Substring(newName.Length - 1, 1) == "_")
                            {
                                newName = newName.Substring(0, newName.Length - 1);
                            }

                            //Check for name already exists

                            string newNamewithcount = "";
                            if (instances.Any(instance => instance.Name == newName))
                            {
                                newNamewithcount = newName + "_" + renamecounter.ToString("D3");
                                while (instances.Any(instance => instance.Name == newNamewithcount))
                                {

                                    renamecounter++;
                                    newNamewithcount = newName + "_" + renamecounter.ToString("D3");
                                }
                            }
                            else
                            {
                                newNamewithcount = newName;
                            }
                                                        
                           
                            tdi.Name = newNamewithcount;
                            //recurssive function to check name
                            //string finalname = RecursiveSearch(ref newName, ref ts);
                        }
                        catch (Exception)
                        {
                            // Reset to old name
                            tdi.Name = oldName;
                            errorCount++;
                        }
                    }

                    if (errorCount > 0)
                    {
                        MessageBox.Show(errorCount.ToString() + " instances could not be renamed.");
                    }
                }
            }

            return null;
        }


        public override string Name
        {
            get { return "Rename Attributes"; }
        }

    }

    public class PopulateInstanceDescription : TCAddOnTask
    {
        public override Type ApplicableType
        {
            get { return typeof(TestSheet); }
        }

        public override TCObject Execute(TCObject objectToExecuteOn, TCAddOnTaskContext taskContext)
        {

            if (objectToExecuteOn is TestSheet)
            {
                TestSheet ts = (TestSheet)objectToExecuteOn;

                IEnumerable<TDElementWrapper> elementWrappers = TestsheetElementsHelper.GetElementsRecursively(ts);

                var instances = ts.Instances.Items.ToArray();
                int errorCount = 0;
                foreach (TDInstance tdi in ts.Instances.Items)
                {
                    String desc = "";

                    try
                    {

                        for (int i = 0; i < elementWrappers.Count(); i++)
                        {
                            TDElementWrapper ewvar = elementWrappers.ElementAt(i);
                            if (ewvar != null)
                            {
                                //tdi.Name += ElementToValueMapper.GetInstanceValueStringForElement(tdi, ew);
                                desc += ElementToValueMapper.GetInstanceValueStringForElement(tdi, ewvar) + "_";
                            }
                        }

                        //Loop to make sure that there are not continuous _                            
                        while (desc.Contains("__"))
                        {
                            desc = desc.Replace("__", "_");
                        }

                        //Check for the trailing character and omit it
                        if (desc.Substring(desc.Length - 1, 1) == "_")
                        {
                            desc = desc.Substring(0, desc.Length - 1);
                        }

                        tdi.Description = desc;

                    }
                    catch (Exception)
                    {
                        // Reset to empty description
                        tdi.Description = "";
                        errorCount++;
                    }
                }

                if (errorCount > 0)
                {
                    MessageBox.Show(errorCount.ToString() + " instance's description could not be renamed.");
                }
            }


            return null;
        }


        public override string Name
        {
            get { return "Populate Instance Description"; }
        }

    }

    
    internal class TDInstanceNameComponent
    {
        private String _text;
        public String Text
        {
            get { return _text; }
            set { _text = value; }
        }

        private Boolean _isPath = false;
        public Boolean IsPath
        {
            get { return _isPath; }
            set { _isPath = value; }
        }

        public TDInstanceNameComponent(String text)
        {
            _text = text;
        }

        public TDInstanceNameComponent(String text, Boolean isPath)
            : this(text)
        {
            _isPath = isPath;
        }
    }

    internal static class TestsheetElementsHelper
    {
        public static List<TDInstanceNameComponent> Initialise(String rawFormat)
        {
            List<TDInstanceNameComponent> components = new List<TDInstanceNameComponent>();

            String[] split = rawFormat.Split(']');
            foreach (String s in split)
            {
                if (s.Contains('['))
                {
                    String[] subS = s.Split('[');

                    components.Add(new TDInstanceNameComponent(subS[0]));
                    components.Add(new TDInstanceNameComponent(subS[1], true));
                }
                else
                {
                    components.Add(new TDInstanceNameComponent(s));
                }
            }

            return components;
        }

        public static IEnumerable<TDElementWrapper> GetElementsRecursively(TestSheet testSheet)
        {
            return GetElementsRecursively(testSheet.Items.Select(a => new TDElementWrapper { TDElement = a, ParentWrapper = null }));
        }

        public static IEnumerable<TDElementWrapper> GetAttrRecursively(TDAttribute tdattribute)
        {
            return GetElementsRecursively(tdattribute.Items.Select(a => new TDElementWrapper { TDElement = a, ParentWrapper = null }));
        }

        private static IEnumerable<TDElementWrapper> GetElementsRecursively(IEnumerable<TDElementWrapper> elements)
        {
            foreach (TDElementWrapper parentElement in elements)
            {
                yield return parentElement;

                IEnumerable<TDElementWrapper> children = GetElementsRecursively(parentElement.TDElement.Items.Select(a => new TDElementWrapper { TDElement = a, ParentWrapper = parentElement }));
                foreach (TDElementWrapper child in children)
                {
                    yield return child;
                }

                TDClass referencedClass = (parentElement.TDElement as TDAttribute) == null ? null : (parentElement.TDElement as TDAttribute).ReferencedClass;
                if (referencedClass != null)
                {
                    IEnumerable<TDElementWrapper> childrenInClass = GetElementsRecursively(referencedClass.Attributes.Select(a => new TDElementWrapper { TDElement = a, ParentWrapper = parentElement }));
                    foreach (TDElementWrapper child in childrenInClass)
                    {
                        yield return child;
                    }
                }
            }
        }

    }

    internal static class ExcelHelper
    {
        public static DataTable ReadExcelFile(String path, String sheetName)
        {
            OleDbConnection conn = null;
            DataTable dt = new DataTable();

            try
            {
                conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path
                                                + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\"");
                OleDbCommand cmd = new OleDbCommand("Select * from [" + sheetName + "$]", conn);
                cmd.CommandType = CommandType.Text;
                OleDbDataAdapter adapt = new OleDbDataAdapter(cmd);

                conn.Open();
                adapt.Fill(dt);
            }
            catch (Exception ex)
            {
                throw new Exception("Error reading Excel file", ex);
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

            return dt;
        }

    }
}
