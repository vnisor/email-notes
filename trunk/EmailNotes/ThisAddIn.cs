using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;

namespace EmailNotes
{
    public partial class ThisAddIn
    {

        private Office.CommandBarButton btn;
        private Outlook.MailItem mail;
        // the id of the field we are working with
        private const string Notes_CustomNotes = "Notes.Email.Custom";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.ItemContextMenuDisplay += new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(
                Application_ItemContextMenuDisplay);
        }

        void Application_ItemContextMenuDisplay(Office.CommandBar CommandBar, Outlook.Selection Selection)
        {
            if (Selection.Count > 0)
            {
                mail = Selection[1] as Outlook.MailItem;
                if (mail != null)
                {
                    btn = CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, missing) as Office.CommandBarButton;
                    btn.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIconAndCaption;
                    btn.Caption = "Email Notes...";
                    btn.FaceId = 1996;
                    btn.Click += new Office._CommandBarButtonEvents_ClickEventHandler(btn_EmailNotes);
                }
            }
        }
        public static UserProperty GetCustomNotestUserProperty(MailItem mailItem)
        {
            
            {
                UserProperty _userProperty = mailItem.UserProperties
                    .Find(Notes_CustomNotes, true /* search custom fields */);

                if (_userProperty == null)
                {
                    // Add the UP because it does not exist
                    mailItem.UserProperties.Add(
                        Notes_CustomNotes,             // Name
                        OlUserPropertyType.olText,      // Type
                        false,                          // Add it to the folder
                        0);                             // Display Format

                    _userProperty = mailItem.UserProperties
                        .Find(Notes_CustomNotes,
                        true /* search custom fields */);
                    mailItem.Save();
                }

                return _userProperty;
            }
        }


        void btn_EmailNotes(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (mail != null)
            {
//                 string subject = mail.Subject;
//                 string filter = @"@SQL=""urn:schemas:httpmail:subject"" like '%" + subject + "%'";
//                 Outlook.Table tbl = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).GetTable(filter, Outlook.OlTableContents.olUserItems);
//                 string result = "";
//                 while (!tbl.EndOfTable)
//                 {
//                     Outlook.Row row = tbl.GetNextRow();
//                     string EntryID = row["EntryID"].ToString();
//                     Outlook.MailItem oMail = (Outlook.MailItem)Application.Session.GetItemFromID(EntryID, Type.Missing);
//                     result += oMail.Subject + " from " + oMail.SenderName + " on " + oMail.SentOn.ToString() + System.Environment.NewLine;
//                     // TODO: Actually delete it (oMail.Delete())
//                 }


                Outlook.UserProperty oProp = GetCustomNotestUserProperty(mail);
                string NotesStr = oProp.Value.ToString();
                //oProp.Value = "Umar inam";

                NoteItem nItem;// = new NoteItem();
                nItem = Application.CreateItem(OlItemType.olNoteItem) as NoteItem;

                nItem.Body = oProp.Value.ToString();
                nItem.Height = 345;
                nItem.Width = 545;

                nItem.Display(1);
                mail.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x10800003", 771);
                oProp.Value = nItem.Body;
                
                string noteSubject = nItem.Subject;// = Notes_CustomNotes;
                nItem.Delete();
                mail.Save();
                nItem = (Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems).Items.Find("[Subject] ='" + noteSubject + "'")) as NoteItem;
                if (null != nItem)
                    nItem.Delete();
            }
        }


//         private void test()
//         {
//             //////////////////////////////////////////////////////////////////////////
//             {
//             Outlook.ApplicationClass myOlApp = null;
//             Outlook.NameSpace newNS = null;
//             myOlApp = new Outlook.ApplicationClass();
//             newNS = myOlApp.GetNamespace("MAPI");
//             Outlook.MAPIFolder mapi1 = newNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
//             string filter = "";
//             Outlook.Table inboxTable = mapi1.GetTable(filter, Outlook.OlTableContents.olUserItems);
//             while (!inboxTable.EndOfTable)
//             {
//                 // get the row item
//                 Outlook.Row myrow = inboxTable.GetNextRow();
//                 string subject = (string)myrow["Subject"];
//                 string EntryID = "";
//                 EntryID = Convert.ToString(myrow["EntryID"]);
//                 string StoreID = "";
//                 StoreID = Convert.ToString(mapi1.StoreID);
//                 if (subject == "Today's WeatherDirect Forecast for Hensall")
//                 {
//                     Outlook.MailItem oMail = (Outlook.MailItem)myOlApp.Session.GetFolderFromID(EntryID, StoreID);
//                     oMail.Delete();
//                 }
//                 // release objects
//                 myrow = null;
//             }
//             }
// 
//             //////////////////////////////////////////////////////////////////////////
//         }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
