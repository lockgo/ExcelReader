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
        List<customers> CustomerHistoryList = new List<customers>();
        int totalAmountOfCustomers = 0;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialogIndex.ShowDialog();
            String path = openFileDialogIndex.FileName;
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
                ContactBox.Text = (ws.Cells[3, 6]).Value2.ToString();
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
            DialogResult result = openFileDialogIndex.ShowDialog();
            String path = openFileDialogIndex.FileName;

            if (openFileDialogIndex.FileName != "")
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(path);
                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(2);
                //List<Indexcustomers> CustomerList = new List<Indexcustomers>();

                int GuestRow = 0;
                int sheetCounter = 2;
                int oldWounds = 0;
                while (sheetCounter < wb.Worksheets.Count)
                {

                    while (((ws.Cells[GuestRow + 3, 1]).Value2 != null))
                    {
                        Indexcustomers CustomerTemp = new Indexcustomers();
                        customers CustomerHistoryTemp = new customers();
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
                            ContactBox.Text = (ws.Cells[GuestRow + 3, 4]).Value2.ToString();
                        }
                        else
                        {
                            ContactBox.Text = "-NA-";
                        }
                        CustomerTemp.GuestNumber = guestNumberBox.Text;
                        CustomerTemp.CasioNumber = clubNumberBox.Text;
                        CustomerTemp.Guest = GuestNameBox.Text;
                        CustomerTemp.Contact = ContactBox.Text;

                        ////////////////////////////////////////////
                        CustomerHistoryTemp.ClubNumber = guestNumberBox.Text;
                        CustomerHistoryTemp.CasioNumber = clubNumberBox.Text;
                        CustomerHistoryTemp.Guest = GuestNameBox.Text;
                        CustomerHistoryTemp.Contact = ContactBox.Text;
                        ////////////////////////////////////////////
                        CustomerList.Add(CustomerTemp);//I might merge these two objects in the future. It was needed to have them as different objects at the time.
                        CustomerHistoryList.Add(CustomerHistoryTemp);//I might merge these two objects in the future. It was needed to have them as different objects at the time.
                        listNames.Items.Add(CustomerList[oldWounds + GuestRow].GuestNumber + " \t" + CustomerList[oldWounds + GuestRow].CasioNumber + " \t" + CustomerList[oldWounds + GuestRow].Guest + " \t" + CustomerList[oldWounds + GuestRow].Contact);
                        GuestRow++;
                    }
                    oldWounds = oldWounds + GuestRow;
                    GuestRow = 0;
                    sheetCounter++;
                    ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.get_Item(sheetCounter);
                }




                textBox1.Text = oldWounds.ToString();
                totalAmountOfCustomers = oldWounds;
                wb.Close();
                excel.Quit();
                MessageBox.Show("Completed ");
                excel = null;
            }

        }

        private void listNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            Indexcustomers load_customer;
            customers load_customerHistory;
            if (listNames.SelectedIndex >= 0)// && listNames.SelectedIndex <= Pathlist.Capacity)//somehow SelectedIndex could be less then 0.
            {
                load_customer = CustomerList[listNames.SelectedIndex];
                load_customerHistory = CustomerHistoryList[listNames.SelectedIndex];
            }
            else
            {
                load_customer = CustomerList[0];
                load_customerHistory = CustomerHistoryList[0];
            }

            //richTextBox1.Text = CustomerHistoryList[listNames.SelectedIndex].CasioNumber; //debug test


            guestNumberBox.Text = load_customer.GuestNumber;
            clubNumberBox.Text = load_customer.CasioNumber;
            GuestNameBox.Text = load_customer.Guest;
            ContactBox.Text = load_customer.Contact;

            EventlistBox.DataSource = load_customerHistory.Events;

            //DateBox.Text = load_customerHistory.Guest;

        }

        private void addButton_Click(object sender, EventArgs e)
        {
            int selectedNameOfList = 0;
            if (listNames.SelectedIndex < 0)
            {
                selectedNameOfList = 0;
            }
            else
            {
                selectedNameOfList = listNames.SelectedIndex;
            }
            //customers are added to customerHistory;
            CustomerHistoryList[selectedNameOfList].historyCount = CustomerHistoryList[selectedNameOfList].historyCount + 1;
            textBox1.Text = CustomerHistoryList[selectedNameOfList].ClubNumber;
            if (DateBox.Text != "")
            {
                CustomerHistoryList[selectedNameOfList].Dates.Add(DateBox.Text);
            }
            else
            {
                CustomerHistoryList[selectedNameOfList].Dates.Add("NA");
            }

            if (DescriptionBox.Text != "")
            {
                CustomerHistoryList[selectedNameOfList].Description.Add(DescriptionBox.Text);
            }
            else
            {
                CustomerHistoryList[selectedNameOfList].Description.Add("NA");
            }

            if (InBox.Text != "")
            {
                CustomerHistoryList[selectedNameOfList].moneyIN.Add(Convert.ToDouble(InBox.Text));
            }
            else
            {
                CustomerHistoryList[selectedNameOfList].moneyIN.Add(0);
            }

            if (OutBox.Text != "")
            {
                CustomerHistoryList[selectedNameOfList].moneyOUT.Add(Convert.ToDouble(OutBox.Text));
            }
            else
            {
                CustomerHistoryList[selectedNameOfList].moneyOUT.Add(0);
            }

            double inMoney = Convert.ToDouble(InBox.Text);
            double outMoney = Convert.ToDouble(OutBox.Text);

            //CustomerHistoryList[selectedNameOfList].Balance.Add(Convert.ToDouble(BalanceBox.Text));
            CustomerHistoryList[selectedNameOfList].Balance.Add(inMoney - outMoney);
            BalanceBox.Text = (inMoney - outMoney).ToString();

            if(RemarkBox.Text != "")
            {
                CustomerHistoryList[selectedNameOfList].Remarks.Add(RemarkBox.Text);
            }
            else
            {
                CustomerHistoryList[selectedNameOfList].Remarks.Add("NA");
            }
            CustomerHistoryList[selectedNameOfList].Events.Add("Date: " + DateBox.Text + " Description: " + DescriptionBox.Text + " In: " + InBox.Text + " Out: " + OutBox.Text + " Balance: " + BalanceBox.Text + " Remark: " + RemarkBox.Text + " ");
            //EventlistBox.Refresh();
            EventlistBox.DataSource = listNames.DataSource;//Refresh doesnot work, this is a work around to see updates.
            EventlistBox.DataSource = CustomerHistoryList[selectedNameOfList].Events;
        }

        private void editButton_Click(object sender, EventArgs e)
        {
            int listNamesIndex = 0;
            int EventListBoxIndex = 0;
            if (listNames.SelectedIndex >= 0 && EventlistBox.SelectedIndex >= 0)
            {
                listNamesIndex = listNames.SelectedIndex;
                EventListBoxIndex = EventlistBox.SelectedIndex;
            }
            else
            {
                if (listNames.SelectedIndex < 0 && EventlistBox.SelectedIndex >= 0)
                {
                    listNamesIndex = 0;
                    EventListBoxIndex = EventlistBox.SelectedIndex;
                }
                else if (listNames.SelectedIndex >= 0 && EventlistBox.SelectedIndex < 0)
                {
                    listNamesIndex = listNames.SelectedIndex;
                    EventListBoxIndex = 0;
                }
                else
                {
                    listNamesIndex = 0;
                    EventListBoxIndex = 0;
                }
            }


            if (DateBox.Text != "")
            {
                CustomerHistoryList[listNamesIndex].Dates[EventListBoxIndex] = DateBox.Text;
            }
            else
            {
                CustomerHistoryList[listNamesIndex].Dates[EventListBoxIndex] = "NA";
            }

            if (DescriptionBox.Text != "")
            {
                CustomerHistoryList[listNamesIndex].Description[EventListBoxIndex] = DescriptionBox.Text;
            }
            else
            {
                CustomerHistoryList[listNamesIndex].Description[EventListBoxIndex] = ("NA");
            }

            if (InBox.Text != "")
            {
                CustomerHistoryList[listNamesIndex].moneyIN[EventListBoxIndex] = (Convert.ToDouble(InBox.Text));
            }
            else
            {
                CustomerHistoryList[listNamesIndex].moneyIN[EventListBoxIndex] = (0);
            }

            if (OutBox.Text != "")
            {
                CustomerHistoryList[listNamesIndex].moneyOUT[EventListBoxIndex] = (Convert.ToDouble(OutBox.Text));
            }
            else
            {
                CustomerHistoryList[listNamesIndex].moneyOUT[EventListBoxIndex] = (0);
            }

            //CustomerHistoryList[listNamesIndex].Balance[EventListBoxIndex] = Convert.ToDouble(BalanceBox.Text);
            CustomerHistoryList[listNamesIndex].Balance[EventListBoxIndex] = Convert.ToDouble(InBox.Text) - Convert.ToDouble(OutBox.Text);

            if (RemarkBox.Text != "")
            {
                CustomerHistoryList[listNamesIndex].Remarks.Add(RemarkBox.Text);
            }
            else
            {
                CustomerHistoryList[listNamesIndex].Remarks.Add("NA");
            }
            CustomerHistoryList[listNamesIndex].Events[EventListBoxIndex] = ("Date: " + DateBox.Text + " Description: " + DescriptionBox.Text + " In: " + InBox.Text + " Out: " + OutBox.Text + " Balance: " + BalanceBox.Text + " Remark: " + RemarkBox.Text + " ");

            EventlistBox.DataSource = listNames.DataSource;//Refresh doesnot work, this is a work around to see updates.
            EventlistBox.DataSource = CustomerHistoryList[listNames.SelectedIndex].Events;
            EventlistBox.Refresh();


        }

        private void deleteButton_Click(object sender, EventArgs e)//Customer did not actually want anything removed, wanted to know when things where deleted.
        {
            int listNamesIndex = 0;
            int EventListBoxIndex = 0;
            if (listNames.SelectedIndex >= 0 && EventlistBox.SelectedIndex >= 0)
            {
                listNamesIndex = listNames.SelectedIndex;
                EventListBoxIndex = EventlistBox.SelectedIndex;
            }
            else
            {
                if (listNames.SelectedIndex < 0 && EventlistBox.SelectedIndex >= 0)
                {
                    listNamesIndex = 0;
                    EventListBoxIndex = EventlistBox.SelectedIndex;
                }
                else if (listNames.SelectedIndex >= 0 && EventlistBox.SelectedIndex < 0)
                {
                    listNamesIndex = listNames.SelectedIndex;
                    EventListBoxIndex = 0;
                }
                else
                {
                    listNamesIndex = 0;
                    EventListBoxIndex = 0;
                }

                CustomerHistoryList[listNamesIndex].Dates[EventListBoxIndex] = "-NA-";
                CustomerHistoryList[listNamesIndex].Description[EventListBoxIndex] = "-NA-";
                CustomerHistoryList[listNamesIndex].moneyIN[EventListBoxIndex] = 0;
                CustomerHistoryList[listNamesIndex].moneyOUT[EventListBoxIndex] = 0;
                CustomerHistoryList[listNamesIndex].Balance[EventListBoxIndex] = 0;
                CustomerHistoryList[listNamesIndex].Remarks[EventListBoxIndex] = "-NA-";
                CustomerHistoryList[listNamesIndex].Events[EventListBoxIndex] = "-Deleted-";

                EventlistBox.DataSource = listNames.DataSource;//Refresh doesnot work, this is a work around to see updates.
                EventlistBox.DataSource = CustomerHistoryList[listNames.SelectedIndex].Events;
                EventlistBox.Refresh();
            }


        }

        private void EventlistBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            int listNamesIndex = 0;
            int EventListBoxIndex = 0;
            if (listNames.SelectedIndex >= 0 && EventlistBox.SelectedIndex >= 0)
            {
                listNamesIndex = listNames.SelectedIndex;
                EventListBoxIndex = EventlistBox.SelectedIndex;
            }
            else
            {
                if (listNames.SelectedIndex < 0 && EventlistBox.SelectedIndex >= 0)
                {
                    listNamesIndex = 0;
                    EventListBoxIndex = EventlistBox.SelectedIndex;
                }
                else if (listNames.SelectedIndex >= 0 && EventlistBox.SelectedIndex < 0)
                {
                    listNamesIndex = listNames.SelectedIndex;
                    EventListBoxIndex = 0;
                }
                else
                {
                    listNamesIndex = 0;
                    EventListBoxIndex = 0;
                }
            }
            DateBox.Text = CustomerHistoryList[listNamesIndex].Dates[EventListBoxIndex];
            DescriptionBox.Text = CustomerHistoryList[listNamesIndex].Description[EventListBoxIndex];
            InBox.Text = (CustomerHistoryList[listNamesIndex].moneyIN[EventListBoxIndex]).ToString();
            OutBox.Text = (CustomerHistoryList[listNamesIndex].moneyOUT[EventListBoxIndex]).ToString();
            BalanceBox.Text = (CustomerHistoryList[listNamesIndex].Balance[EventListBoxIndex]).ToString();
            RemarkBox.Text = CustomerHistoryList[listNamesIndex].Remarks[EventListBoxIndex].ToString();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = saveFileDialog1.ShowDialog();
            String path = saveFileDialog1.FileName;
            //String path = openFileDialogIndex.FileName;

            if (saveFileDialog1.FileName != "")
            {
                textBox1.Text = NewExcelFile(path, CustomerList, CustomerHistoryList, totalAmountOfCustomers);
            }

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            if (MessageBox.Show("Are you sure you want to Quit.", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                //e.cancel = true;
                Application.Exit();
                Environment.Exit(0);

            }
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to make a new file?.", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) ==
                DialogResult.Yes)
            {
                List<Indexcustomers> CustomerList = new List<Indexcustomers>();
                List<customers> CustomerHistoryList = new List<customers>();
                totalAmountOfCustomers = 0;
                listNames = new ListBox();
                EventlistBox = new ListBox();
                listNames.DataSource = CustomerList;
                EventlistBox.DataSource = CustomerHistoryList;

            }
        }

        public static String NewExcelFile(string filePath, List<Indexcustomers> CustomerList, List<customers> CustomerHistoryList, int totalNumber)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            for (int numberOfCustomers = totalNumber-1; numberOfCustomers >= 0; numberOfCustomers--)
            {
                xlWorkBook.Worksheets.Add();

                Microsoft.Office.Interop.Excel.Worksheet worksheet = xlApp.Worksheets["Sheet" + ((totalNumber - numberOfCustomers + 3).ToString())];
                worksheet.Name = "Guest " + (CustomerHistoryList[numberOfCustomers].ClubNumber).ToString();
            }

            for (int numberOfCustomers = 0; numberOfCustomers < totalNumber; numberOfCustomers++)
            {
                
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(numberOfCustomers+1);
                
                //xlWorkSheet.Cells[Y, X] 
                //xlWorkSheet.Cells[1, 1] = "http://csharp.net-informations.com";

                xlWorkSheet.Cells[1, 1] = "Guest:";
                xlWorkSheet.Cells[1, 2] = CustomerHistoryList[numberOfCustomers].Guest;
                xlWorkSheet.Cells[1, 5] = "Club Number: ";
                xlWorkSheet.Columns["E:E"].ColumnWidth = 12.57;

                
                xlWorkSheet.Cells[1, 6] = CustomerHistoryList[numberOfCustomers].ClubNumber;

                xlWorkSheet.Cells[3, 1] = "Contact:";
                xlWorkSheet.Cells[3, 2] = CustomerHistoryList[numberOfCustomers].Contact;
                xlWorkSheet.Cells[3, 5] = "Casio Number:";
                xlWorkSheet.Cells[3, 6] = CustomerHistoryList[numberOfCustomers].CasioNumber;


                xlWorkSheet.Cells[5, 1] = "Date";
                xlWorkSheet.Cells[5, 2] = "Description";
                xlWorkSheet.Cells[5, 3] = "In";
                xlWorkSheet.Cells[5, 4] = "Out";
                xlWorkSheet.Cells[5, 5] = "Balance";
                xlWorkSheet.Cells[5, 6] = "Remark";

                for (int i = 0; i < CustomerHistoryList[numberOfCustomers].historyCount; i++)
                {
                    xlWorkSheet.Cells[i+6, 1] = CustomerHistoryList[numberOfCustomers].Dates[i];
                    xlWorkSheet.Cells[i+6, 2] = CustomerHistoryList[numberOfCustomers].Description[i];
                    xlWorkSheet.Cells[i+6, 3] = CustomerHistoryList[numberOfCustomers].moneyIN[i];
                    xlWorkSheet.Cells[i+6, 4] = CustomerHistoryList[numberOfCustomers].moneyOUT[i];
                    xlWorkSheet.Cells[i+6, 5] = CustomerHistoryList[numberOfCustomers].Balance[i];
                    xlWorkSheet.Cells[i+6, 6] = CustomerHistoryList[numberOfCustomers].Remarks[i];
                }
            }

            xlWorkBook.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            MessageBox.Show("Completed ");
            xlApp = null;

            return "something";
        }
    }
}
