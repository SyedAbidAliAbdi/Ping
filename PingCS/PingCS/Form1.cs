using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using XLS = Microsoft.Office.Interop.Excel;

namespace PingCS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public bool CheckMe(string St)
        {
        var ping = new Ping();
        var reply = ping.Send("127.0.0.1", 60 * 300); // 1 minute time out (in ms)
            if (reply.Status == IPStatus.Success)
                {
                    return true;
                }
                else
                {
                    return false;
                }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //var i =  ExcelFileCheck(Application.StartupPath  & "\\Files\\book1.xls", 5, 3);
            if (ExcelFileCheck(Application.StartupPath + "\\Files\\book1.xls", 5, 3) == 0.0)
            {
                using (System.IO.StreamReader reader = new System.IO.StreamReader(@"Files\ServerIP.txt"))
                {
                    while (reader.Peek() >= 0)
                    {
                        // check system is live or not
                        string ServerIP = reader.ReadLine();

                        if (CheckMe(ServerIP) == true)
                        {
                            MessageBox.Show(ServerIP + " Server is On");
                            // wait for one minutes
                            System.Diagnostics.ProcessStartInfo thePSI = new System.Diagnostics.ProcessStartInfo("shutdown");

                            thePSI.Arguments = "/m \\\\" + ServerIP + " /s /t 01";

                            System.Diagnostics.Process.Start(thePSI);
                            System.Threading.Thread.Sleep(10000);   
                        }
                        else
                        {
                            MessageBox.Show(ServerIP + " Server is Down");
                        }
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            XLS.Application myapp = new XLS.Application();

            XLS.Workbook wb = myapp.Workbooks.Open(@"C:\\Users\\abid.s.1\\Desktop\\book2.xlsx");

            XLS.Worksheet sheet = (XLS.Worksheet)wb.Worksheets.get_Item(1);

            var cell = (XLS.Range)sheet.Cells[3, 5];

            double cellval = cell.Value2;
            
            MessageBox.Show ("Hello");

            
            myapp.DisplayAlerts = false;
            
            myapp.Workbooks.Close();
            
            myapp.Quit();
	        
        }

        public double ExcelFileCheck(String FilePath, int X, int Y)
        {
            XLS.Application myapp = new XLS.Application();

            XLS.Workbook wb = myapp.Workbooks.Open(@FilePath);

            XLS.Worksheet sheet = (XLS.Worksheet)wb.Worksheets.get_Item(1);

            var cell = (XLS.Range)sheet.Cells[Y, X];

            return cell.Value2;
                        
            myapp.DisplayAlerts = false;
            
            myapp.Workbooks.Close();
            
            myapp.Quit();
        }
    }
}
