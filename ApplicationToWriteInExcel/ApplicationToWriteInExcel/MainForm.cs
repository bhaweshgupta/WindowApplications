/*
 * Created by SharpDevelop.
 * User: 765454
 * Date: 4/5/2019
 * Time: 4:05 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel; 

namespace ApplicationToWriteInExcel
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		static List<Employee> Employees=new List<Employee>();
			public static int x=10;
			
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		void Button2_Click(object sender, EventArgs e)
		{

							
					Employee emp=new Employee();
					emp.Name=textBox1.Text;
					emp.ID=Convert.ToInt32(textBox2.Text);
					emp.Age=Convert.ToInt32(textBox3.Text);
					//Employees=
					Employees.Add(emp);		
			
			
		}
		
		
		void Button1_Click(object sender, EventArgs e)
		{
			Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            
           
             xlWorkSheet.Cells[1, 1] = "Name";
             xlWorkSheet.Cells[1, 2] = "ID";
             xlWorkSheet.Cells[1, 3] = "Age";
             int i=2;
             foreach(var emp in Employees)
             {
             	xlWorkSheet.Cells[i,1]=emp.Name;
             	xlWorkSheet.Cells[i, 2] =emp.ID;
             	xlWorkSheet.Cells[i,3] =emp.Age;
             	i++;
             }


            xlWorkBook.SaveAs("d:\\sample.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
		}
		
		void Button3_Click(object sender, EventArgs e)
		{
			foreach(var Emp in Employees)
				listBox1.Items.Add(Emp.Name+" "+Emp.Age+" "+Emp.ID);
		}
	}
}
