using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace OutlookAddIn1
{
    partial class VodaMeeting
    {
        
        #region Form Region Factory 
        // [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)]
        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass("IPM.Appointment.Contoso")]
        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Appointment)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("OutlookAddIn1.FormRegion1")]
        public partial class FormRegion1Factory
        {
            // Occurs before the form region is initialized.
            // To prevent the form region from appearing, set e.Cancel to true.
            // Use e.OutlookItem to get a reference to the current Outlook item.
            private void FormRegion1Factory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
                
            }
        }

        #endregion

        // Occurs before the form region is displayed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
              
        private void FormRegion1_FormRegionShowing(object sender, System.EventArgs e)
        {              
           Outlook.AppointmentItem myItem = (Outlook.AppointmentItem)this.OutlookItem;
           if (myItem.Organizer != null)
            {
              this.Visible = false;
              this.Size = new System.Drawing.Size(0, 0);
            }                       
            //myItem.Forward += new Outlook.ItemEvents_10_ForwardEventHandler(MyMailItem_Open);            
        }
       
        //private void MyMailItem_Open(object Response, ref bool Cancel)
        //{
          //  Outlook.AppointmentItem myItem = (Outlook.AppointmentItem)this.OutlookItem;
            //if (myItem.Organizer != myItem.Session.CurrentUser.Name) {
              //  MessageBox.Show("You can not forward this activity because you are not owner");
                //Cancel = true;
            //}            
       // }
            // Occurs when the form region is closed.
            // Use this.OutlookItem to get a reference to the current Outlook item.
            // Use this.OutlookFormRegion to get a reference to the form region.
            private void FormRegion1_FormRegionClosed(object sender, System.EventArgs e)
        {
            this.Visible = true;
            this.Size = new System.Drawing.Size(887, 451);
        }        
        private void BtnImport_Click(object sender, EventArgs e)
        {
            Outlook.AppointmentItem oAppointment = (Outlook.AppointmentItem)this.OutlookItem;
            //oAppointment.Body = "PROJECT Name :  " + textBox1.Text;
            //oAppointment.Body = oAppointment.Body + "PROJECT Goal :  " + richTextBox1.Text;
            //oAppointment.Body = oAppointment.Body + "PROJECT Agenda :  " + richTextBox2.Text;
            StringBuilder rtf = new StringBuilder();
            string Location = oAppointment.Location;
            string Subject = oAppointment.Subject;

            ////// Header ///////////////
            rtf.Append(@"{\rtf1\ansi\deff0 {\fonttbl {\f0 Arial;}}");                
            rtf.Append(@"\trowd\trgaph144");
            rtf.Append(@"\clbrdrt\brdrengrave\clbrdrl\brdrengrave\clbrdrb\brdremboss\clbrdrr\brdremboss");
            rtf.Append(@"\cellx10000");
            rtf.Append(@"DRAFT\intbl\cell");
            rtf.Append(@"\row");
            ////// Project Textbox ///////////////
            rtf.Append(@"{\rtf1\ansi\deff0 {\fonttbl {\f0 Arial;}}");
            rtf.Append(@"\trowd\trgaph144");
            rtf.Append(@"\clbrdrt\brdrengrave\clbrdrl\brdrengrave\clbrdrb\brdremboss\clbrdrr\brdremboss");
            rtf.Append(@"\cellx1000");            
            rtf.Append(@"PROJECT:\intbl\cell");
            rtf.Append(@"\clbrdrt\brdrengrave\clbrdrl\brdrengrave\clbrdrb\brdremboss\clbrdrr\brdremboss");
            rtf.Append(@"\cellx10000");
            rtf.Append(richTextBox3.Rtf);
            rtf.Append(@"\intbl\cell");            
            rtf.Append(@"\row");
            ////// Location Textbox ///////////////
            if (Location != null)
            {
                rtf.Append(@"{\rtf1\ansi\deff0 {\fonttbl {\f0 Arial;}}");
                rtf.Append(@"\trowd\trgaph144");
                rtf.Append(@"\clbrdrt\brdrengrave\clbrdrl\brdrengrave\clbrdrb\brdremboss\clbrdrr\brdremboss");
                rtf.Append(@"\cellx1000");
                rtf.Append(@"LOCATION:\intbl\cell");
                rtf.Append(@"\clbrdrt\brdrengrave\clbrdrl\brdrengrave\clbrdrb\brdremboss\clbrdrr\brdremboss");
                rtf.Append(@"\cellx10000");
                rtf.Append(Location);
                rtf.Append(@"\intbl\cell");
                rtf.Append(@"\row");
            }
            ////// Subject Textbox ///////////////
            if (Subject != null)
            {
                rtf.Append(@"{\rtf1\ansi\deff0 {\fonttbl {\f0 Arial;}}");
                rtf.Append(@"\trowd\trgaph144");
                rtf.Append(@"\clbrdrt\brdrengrave\clbrdrl\brdrengrave\clbrdrb\brdremboss\clbrdrr\brdremboss");
                rtf.Append(@"\cellx1000");
                rtf.Append(@"SUBJECT:\intbl\cell");
                rtf.Append(@"\clbrdrt\brdrengrave\clbrdrl\brdrengrave\clbrdrb\brdremboss\clbrdrr\brdremboss");
                rtf.Append(@"\cellx10000");
                rtf.Append(Subject);
                rtf.Append(@"\intbl\cell");
                rtf.Append(@"\row");
            }
            ////// Goal Textbox ///////////////
            rtf.Append(@"{\rtf1\ansi\deff0 {\fonttbl {\f0 Arial;}}");
            rtf.Append(@"\trowd\trgaph144");
            rtf.Append(@"\clbrdrt\brdrengrave\clbrdrl\brdrengrave\clbrdrb\brdremboss\clbrdrr\brdremboss");
            rtf.Append(@"\cellx1000");
            rtf.Append(@"GOAL:\intbl\cell");
            rtf.Append(@"\clbrdrt\brdrengrave\clbrdrl\brdrengrave\clbrdrb\brdremboss\clbrdrr\brdremboss");
            rtf.Append(@"\cellx10000");
            rtf.Append(richTextBox1.Rtf);
            rtf.Append(@"\intbl\cell");
            rtf.Append(@"\row");
            ////// Agenda Textbox ///////////////
            rtf.Append(@"{\rtf1\ansi\deff0 {\fonttbl {\f0 Arial;}}");
            rtf.Append(@"\trowd\trgaph144");
            rtf.Append(@"\clbrdrt\brdrengrave\clbrdrl\brdrengrave\clbrdrb\brdremboss\clbrdrr\brdremboss");
            rtf.Append(@"\cellx1000");
            rtf.Append(@"AGENDA:\intbl\cell");
            rtf.Append(@"\clbrdrt\brdrengrave\clbrdrl\brdrengrave\clbrdrb\brdremboss\clbrdrr\brdremboss");
            rtf.Append(@"\cellx10000");
            rtf.Append(richTextBox2.Rtf);
            rtf.Append(@"\intbl\cell");
            rtf.Append(@"\row");
            ////// Last //////
            rtf.Append(@"}");
            
            oAppointment.RTFBody = System.Text.Encoding.GetEncoding("iso-8859-9").GetBytes(rtf.ToString());                        
        }       
        private void ButtonLocation(object sender, EventArgs e)
        {
            int y = this.Size.Height - this.BtnImport.Height - 5;
            int x = this.Size.Width - this.BtnImport.Width -5;
            int x1 = this.Size.Width - 150;
            this.BtnImport.Location = new System.Drawing.Point(x, y);
            this.pictureBox1.Location = new System.Drawing.Point(x, 5);
            this.richTextBox1.Size = new System.Drawing.Size(x1, 85);
            this.richTextBox2.Size = new System.Drawing.Size(x1, 120);
        }
    }
}
