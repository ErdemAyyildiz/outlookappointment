using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;


namespace OutlookAddIn1
{    
    public partial class outlookappointment
    {

        Outlook.Inspectors inspectors;
        Outlook.Explorer currentExplorer = null;
                     
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {            
               inspectors = this.Application.Inspectors;
               inspectors.NewInspector +=
               new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            //currentExplorer = this.Application.ActiveExplorer();
            //currentExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(CurrentExplorer_Event);
        }
        private void CurrentExplorer_Event() {
            if (this.Application.ActiveExplorer().Selection.Count > 0) {
                Object selObject = this.Application.ActiveExplorer().Selection[1];
                if (selObject is Outlook.AppointmentItem)
                {
                    Outlook.AppointmentItem apptItem = (selObject as Outlook.AppointmentItem);
                    if (apptItem.Organizer != apptItem.Session.CurrentUser.Name)
                    {
                        //MessageBox.Show("You can not forward this activity because you are not owner");
                        
                    }
                    //inspectors = this.Application.Inspectors;                    
                    //inspectors.NewInspector +=
                    //new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);                    

                }
                else if (selObject is Outlook.MeetingItem)
                {
                    Outlook.MeetingItem apptItem = (selObject as Outlook.MeetingItem);
                    inspectors = this.Application.Inspectors;
                    apptItem.Actions["Forward"].Enabled = false;
                    //inspectors.NewInspector +=
                    //new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
                }
            }                
        }
        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.AppointmentItem mailItem = Inspector.CurrentItem as Outlook.AppointmentItem;
                                        
            if (mailItem != null)
            {                
                mailItem.Forward += new Outlook.ItemEvents_10_ForwardEventHandler(MyMailItem_Open);
                //mailItem.ForwardAsVcal.fo += new Outlook.ApplicationEvents_StartupEventHandler(void(MyMailItem_Open));
            }
            void MyMailItem_Open(object Forward, ref bool Cancel)
            {                
                if (mailItem.Organizer != mailItem.Session.CurrentUser.Name)
                {                    
                    MessageBox.Show("You can not forward this activity because you are not owner");
                    Cancel = true;
                }
            }
        }
     

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);            
        }
        
        #endregion
    }
}
