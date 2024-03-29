﻿using System;
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

        private Office.CommandBarButton btnNotes;
        private Office.CommandBarButton btnID;
        private Outlook.MailItem mail;
        // the id of the field we are working with
        private const string Notes_CustomNotes = "Notes.Email.Custom";
        private const string CatagoryWithNotes = "WithNotes";

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
                    btnID = CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, missing) as Office.CommandBarButton;
                    btnID.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIconAndCaption;
                    btnID.Caption = "Copy Email ID";
                    btnID.FaceId = 224;
                    btnID.Click += new Office._CommandBarButtonEvents_ClickEventHandler(btn_CopyEmailID);

                    btnNotes = CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, missing) as Office.CommandBarButton;
                    btnNotes.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIconAndCaption;
                    btnNotes.Caption = "Email Notes...";
                    btnNotes.FaceId = 1996;
                    btnNotes.Click += new Office._CommandBarButtonEvents_ClickEventHandler(btn_EmailNotes);
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
                Outlook.UserProperty oProp = GetCustomNotestUserProperty(mail);
                string NotesStr = oProp.Value.ToString();
                //oProp.Value = "Umar inam";

                NoteItem nItem;// = new NoteItem();
                nItem = Application.CreateItem(OlItemType.olNoteItem) as NoteItem;

                nItem.Body = oProp.Value.ToString();
                nItem.Height = 345;
                nItem.Width = 545;

                nItem.Display(1);
                //mail.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x10800003", 771);
                if (null != nItem.Body && nItem.Body.Length > 0)
                {
                    oProp.Value = nItem.Body;
                    if (null == mail.Categories || !mail.Categories.Contains(CatagoryWithNotes))
                        mail.Categories = mail.Categories + "," + CatagoryWithNotes;
                }
                else
                {
                    oProp.Value = string.Empty;
                    if (null != mail.Categories && mail.Categories.Contains(CatagoryWithNotes))
                    {
                        mail.Categories = mail.Categories.Replace("," + CatagoryWithNotes, "");
                        mail.Categories = mail.Categories.Replace(CatagoryWithNotes + ",", "");
                        mail.Categories = mail.Categories.Replace(CatagoryWithNotes, "");
                    }

                }
                
                
                string noteSubject = nItem.Subject;// = Notes_CustomNotes;
                nItem.Delete();
                mail.Save();
                nItem = (Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems).Items.Find("[Subject] ='" + noteSubject + "'")) as NoteItem;
                if (null != nItem)
                    nItem.Delete();
            }
        }

        void btn_CopyEmailID(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (mail != null)
            {
                string strEmailID = string.Empty;
                strEmailID = mail.EntryID;
                strEmailID = "[[" + mail.Subject + "|" + "outlook:" + mail.EntryID + "]]";
                Clipboard.SetText(strEmailID);
            }
        }




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
