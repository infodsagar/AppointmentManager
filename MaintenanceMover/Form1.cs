using Microsoft.Office.Interop.Outlook;
using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace MaintenanceMover
{
    public partial class Form1 : Form
    {
        Outlook.Application oApp = null;
        NameSpace mapiNamespace = null;
        MAPIFolder calendarFolder = null;
        Items oCalendarItems = null;
        AppointmentItem oAppoint = null;


        public Form1()
        {
            InitializeComponent();
            oApp = new Outlook.Application();
            mapiNamespace = oApp.GetNamespace("MAPI");
            calendarFolder = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            oCalendarItems = calendarFolder.Items;
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            GetAllCalenderItems();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AddAppointment();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            RemoveAppointment();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            MoveAppointment();
        }

       

        //Get all appointment from calendar
        public void GetAllCalenderItems()
        {

            AppointmentList.Items.Clear();

            try
            {
                foreach (AppointmentItem item in oCalendarItems)
                {

                    Console.WriteLine(item.Subject + " -> " + item.Start.ToLongDateString());
                    AppointmentList.Items.Add(item.Subject + item.Start.ToLongDateString());

                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        // Remove Appointments
        public void RemoveAppointment()
        {

            try
            {
                int i = 0;
                foreach (AppointmentItem item in oCalendarItems)
                {

                     if(i == 0)
                    {
                        item.Delete();
                        i++;
                        item.Save();
                    }
                }
            }
            catch(System.Exception ex)
            {
                MessageBox.Show("The following error occurred: " + ex.Message);
            }
        }
            

        //Add new appointments
        private void AddAppointment()
        {
            oAppoint = (AppointmentItem)
            oApp.CreateItem(OlItemType.olAppointmentItem);
            oAppoint.Start = DateTime.Today.AddDays(6);
            oAppoint.AllDayEvent = true;
            oAppoint.Subject = "Group Project";

            try
            {
                OpenFileDialog attachment = new OpenFileDialog();
                attachment.Multiselect = true;
                attachment.Title = "Select a file to attach";
                

                if(attachment.ShowDialog() == DialogResult.OK && attachment.FileName.Length > 0)
                {
                    foreach (string FileName in attachment.FileNames)
                    {
                        oAppoint.Attachments.Add(FileName);
                    }
                }

                oAppoint.RTFBody = System.Text.Encoding.ASCII.GetBytes("hi");
                oAppoint.Save();
                oAppoint.Display(true);

                
            } 
            catch (System.Exception ex)
            {
                MessageBox.Show("The following error occurred: " + ex.Message);
            }
            
        }

        // Move appointments
        public void MoveAppointment()
        {           

            try
            {
                
                foreach (AppointmentItem item in oCalendarItems)
                {
                    item.Start = dateTimePicker1.Value.Date;
                    item.Save();
                    
                }
            }
            
             catch (System.Exception ex)
            {
                MessageBox.Show("The following error occurred: " + ex.Message);
            }
        }



        // ***********Part of Designer form *************************
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void AppointmentList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }
    }
    }



