using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Diagnostics;
using Outlook2010Sender.Views;
using System.IO;

namespace Outlook2010Sender
{
    public partial class ThisAddIn
    {
        private Office.CommandBar _menuBar;
        private Office.CommandBarPopup _newMenubar;
        public Office.CommandBarButton _buttonOne; // Send emails to selection + template
        public Office.CommandBarButton _buttonTwo; // Send emails to selection new empty template
        public static Outlook.Application app;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            app = this.Application;
            AddMenuBar();
        }
        
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        /// <summary>
        /// Get add-in ribbon extension
        /// </summary>
        /// <returns>Ribbon extension object</returns>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new QMSelectedEmailsBtn();
        }

        public void AddMenuBar()
        {
            _menuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;
            _newMenubar = (Office.CommandBarPopup)_menuBar.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, true);

            if (_newMenubar != null)
            {
                _newMenubar.Caption = "Send Email to selection";

                // Initialize button one
                _buttonOne = (Office.CommandBarButton)_newMenubar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                _buttonOne.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                _buttonOne.Caption = "Email from template";
                _buttonOne.Click += _buttonOne_Click;
                
                // Initialize button two
                _buttonTwo = (Office.CommandBarButton)_newMenubar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 2, true);
                _buttonTwo.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                _buttonTwo.Caption = "New email";
                _buttonTwo.Click += _buttonTwo_Click;

                // menubar visible
                _newMenubar.Visible = true;
            }
        }

        #region Events
        void _buttonTwo_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ThisAddIn.SendMails();
        }

        public void _buttonOne_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            SelectListPopup popup = new SelectListPopup();
            popup.ShowDialog();
        }
        #endregion

        #region static functions
        /// <summary>
        /// Send mail to selected people from message template
        /// </summary>
        /// <param name="templateLocation">Template file (.msg) location</param>
        public static void SendMails(string templateLocation)
        {
            // Check if location is set and its right
            if (templateLocation == null || !File.Exists(templateLocation))
            {
                MessageBox.Show("Invalid template file location: " + templateLocation);
                return;
            }

            // Check if file type is right
            string fType = Path.GetExtension(templateLocation);
            if (fType != ".msg")
            {
                MessageBox.Show("Invalid filetype: " + fType);
                return;
            }

            Outlook.Selection selection = app.ActiveExplorer().Selection;
            Outlook.Folder folder = app.Session.GetDefaultFolder(OlDefaultFolders.olFolderDrafts) as Outlook.Folder;
            Outlook.MailItem mail = app.CreateItemFromTemplate(templateLocation, folder) as Outlook.MailItem;
            
            foreach (Outlook.MailItem item in selection)
            {
                Outlook.Recipient recipent = mail.Recipients.Add(item.Sender.Address);
                recipent.Type = (int)OlMailRecipientType.olBCC;
            }

            mail.Recipients.ResolveAll();
            mail.Display(false);
        }

        /// <summary>
        /// Send new mail to selected people
        /// </summary>
        public static void SendMails()
        {
            Outlook.Selection selection = app.ActiveExplorer().Selection;
            Outlook.Folder folder = app.Session.GetDefaultFolder(OlDefaultFolders.olFolderDrafts) as Outlook.Folder;
            Outlook.MailItem mail = app.CreateItem(OlItemType.olMailItem) as Outlook.MailItem;

            foreach (Outlook.MailItem item in selection)
            {
                Outlook.Recipient recipent = mail.Recipients.Add(item.Sender.Address);
                recipent.Type = (int)OlMailRecipientType.olBCC;
            }

            mail.Recipients.ResolveAll();
            mail.Display(false);
        }
        #endregion

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
