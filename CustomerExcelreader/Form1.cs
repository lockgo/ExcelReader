using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CustomerExcelreader
{
    //using Microsoft.Office.Interop.Excel;
    public partial class Form1 : Form
    {
        List<Indexcustomers> CustomerList = new List<Indexcustomers>();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            String path = openFileDialog1.FileName;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(path);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(1); 
            //ws.Cells[1, 1] = "http://csharp.net-informations.com";

            //textBox1.Text = (ws.Cells[Y, X]).Value2.ToString();

            //textBox1.Text = (ws.Cells[1, 2]).Value2.ToString();

            if ((ws.Cells[1, 2]).Value2 != null)
            {
                guestNumberBox.Text = (ws.Cells[1, 2]).Value2.ToString();
            }
            else
            {
                
            }

            if (ws.Cells[3, 2].Value2 != null)
            {
                GuestNameBox.Text = (ws.Cells[3, 2]).Value2.ToString();
            }
            else
            {
                
            }

            if ((ws.Cells[1, 6]).Value2 != null)
            {
                clubNumberBox.Text = (ws.Cells[1, 6]).Value2.ToString();
            }
            else
            {
                
            }

            if ((ws.Cells[3, 6]).Value2
                != null)
            {
                RemarkBox.Text = (ws.Cells[3, 6]).Value2.ToString();
            }
            else
            {
                
            }

            
            
            
            
            



            wb.Close();
            excel.Quit();
            MessageBox.Show("Completed");
            excel = null;

            

            

        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            String path = openFileDialog1.FileName;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(path);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(2);
            //List<Indexcustomers> CustomerList = new List<Indexcustomers>();

            int GuestRow = 0;
            int sheetCounter = 2;
            
            while (sheetCounter < wb.Worksheets.Count)
            {
                
                while(((ws.Cells[GuestRow + 3, 1]).Value2 != null))
                {
                    Indexcustomers CustomerTemp = new Indexcustomers();
                    if ((ws.Cells[GuestRow + 3, 1]).Value2 != null)
                    {
                        guestNumberBox.Text = (ws.Cells[GuestRow + 3, 1]).Value2.ToString();
                    }
                    else
                    {
                        guestNumberBox.Text = "-NA-";
                    }

                    if ((ws.Cells[GuestRow + 3, 2]).Value2 != null)//Casino number
                    {
                        clubNumberBox.Text = (ws.Cells[GuestRow + 3, 2]).Value2.ToString();
                    }
                    else
                    {
                        clubNumberBox.Text = "-NA-";
                    }

                    if ((ws.Cells[GuestRow + 3, 3]).Value2 != null)//Guest Name
                    {
                        GuestNameBox.Text = (ws.Cells[GuestRow + 3, 3]).Value2.ToString();
                    }
                    else
                    {
                        GuestNameBox.Text = "-NA-";
                    }

                    if ((ws.Cells[GuestRow + 3, 4]).Value2 != null)//remake
                    {
                        RemarkBox.Text = (ws.Cells[GuestRow + 3, 4]).Value2.ToString();
                    }
                    else
                    {
                        RemarkBox.Text = "-NA-";
                    }
                    CustomerTemp.GuestNumber = guestNumberBox.Text;
                    CustomerTemp.CasioNumber = clubNumberBox.Text;
                    CustomerTemp.Guest = GuestNameBox.Text;
                    CustomerTemp.Contact = RemarkBox.Text;
                    CustomerList.Add(CustomerTemp);
                    listNames.Items.Add(CustomerList[GuestRow].GuestNumber + " \t" + CustomerList[GuestRow].CasioNumber + " \t" + CustomerList[GuestRow].Guest + " \t" + CustomerList[GuestRow].Contact);
                    GuestRow++;
                }
                GuestRow = 0;
                sheetCounter++;
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(sheetCounter);
            }




            textBox1.Text = wb.Worksheets.Count.ToString();
            wb.Close();
            excel.Quit();
            MessageBox.Show("Completed ");
            excel = null;

        }

        private void listNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            Indexcustomers load_customer;
            if (listNames.SelectedIndex >= 0)// && listNames.SelectedIndex <= Pathlist.Capacity)//somehow SelectedIndex could be less then 0.
            {
                load_customer = CustomerList[listNames.SelectedIndex];
            }
            else
            {
                load_customer = CustomerList[0];
            }
            guestNumberBox.Text   = load_customer.GuestNumber;
            clubNumberBox.Text  = load_customer.CasioNumber;
            GuestNameBox.Text =   load_customer.Guest;
            RemarkBox.Text = load_customer.Contact;

        }
    }
}
