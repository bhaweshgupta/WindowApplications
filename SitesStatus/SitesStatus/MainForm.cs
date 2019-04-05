/*
 * Created by SharpDevelop.
 * User: 765454
 * Date: 4/1/2019
 * Time: 12:50 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;           
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Windows.Forms;
using System.Net;
namespace SitesStatus
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
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
		
		void Button1_Click(object sender, EventArgs e)
		{
			List<string> mylist=startreadingfromexcel();
			listBox1.DataSource=mylist;
			
			
		}
			public static List<string> startreadingfromexcel()
			{
				List<string> someList =new List<string>();
				Excel.Application xlApp = new Excel.Application();
				Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\765454\Desktop\website.xlsx");
				Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
				Excel.Range xlRange = xlWorksheet.UsedRange;
				int rowCount = xlRange.Rows.Count;
				int colCount = xlRange.Columns.Count;
				for (int i = 2; i <= rowCount; i++)
				{
					object _xVal;
					_xVal= ((Excel.Range)xlWorksheet.Cells[i, 1]).Value2;
					if(xlWorksheet.Cells[i, 1]!=null&&_xVal!=null)
					{
						Uri uriResult;
						string uriName =_xVal.ToString();
						bool result = Uri.TryCreate(uriName, UriKind.Absolute, out uriResult)
							&& uriResult.Scheme == Uri.UriSchemeHttps;
						//System.Console.WriteLine(result);
						if(result==true)
						{
							bool temp=checkifitisup(uriName);
							
							if(temp==true)
							{
								someList.Add("Site is Up");
							}
							else
							{
								someList.Add("Site is down or doest not exist");
							}
								
						}
						else
						{
							someList.Add("Incorrect Format or non Https");
						}
						
					}
				}
				return someList;
			}

			public static bool checkifitisup(string Url)
			{
			string html = string.Empty;
        	string Message = string.Empty;
        	 HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(Url);
        	request.Credentials = System.Net.CredentialCache.DefaultCredentials;
            request.Method = "GET";

            try
            {
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
						using (Stream stream = response.GetResponseStream())
       					using (StreamReader reader = new StreamReader(stream))
						{
							html = reader.ReadToEnd();
						}
                }
            }
            catch (WebException ex)
            {
                Message += ((Message.Length > 0) ? "\n" : "") + ex.Message;
                html="website is down or doesn't exist /n can't fetch data from site ";
            }
			
             return (Message.Length==0);
            
			}
		
		
	}
}
