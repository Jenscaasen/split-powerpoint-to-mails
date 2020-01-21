using System;
using System.Activities.Statements;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;

namespace SendSheetsViaMail
{
    public partial class RibbonMenu
    {
        private void RibbonMenu_Load(object sender, RibbonUIEventArgs e)
        {

        }
        CSVInput form = new CSVInput();
        private void btnSetCsv_Click(object sender, RibbonControlEventArgs e)
        {
            form.Show();
        }

        private void btnSendMails_Click(object sender, RibbonControlEventArgs e)
        {
            form.Show();
            var mappings = form.txtMapping.Text.Split('\n').Skip(1); //skip header
            string subject = form.txtSubject.Text;
            string body = form.txtBody.Text;
            foreach(string map in mappings)
            {
                string sheetRange = map.Split(';')[0];
                string department = map.Split(';')[1];
                string mail = map.Split(';')[2];
                List<int> sheetNumbers = GetSheetNumbers(sheetRange);

                Presentation curCopy = MakeCopyFor(department);
                StripSheets(curCopy, sheetNumbers);
                PrepareMail(mail, department, subject, body);
            }
            form.Hide();

        }

        private void PrepareMail( string mail, string department, string subject, string body)
        {
           Outlook.Application oApp = new Outlook.Application();
            Outlook._MailItem oMailItem = (Outlook._MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
            oMailItem.To = mail;
            oMailItem.Subject = subject;
            oMailItem.Body = body;
            string path = Directory.GetCurrentDirectory();
            oMailItem.Attachments.Add(path + "/" + department + ".pptx");           
            oMailItem.Display(false);
        }

        private void StripSheets(Presentation curCopy, List<int> sheetNumbers)
        {
            List<Slide> doNotDelete = new List<Slide>();
          foreach(int number in sheetNumbers)
            {
                var slide = curCopy.Slides[number];
                doNotDelete.Add(slide);
            }
          for(int i = 1; i < curCopy.Slides.Count + 1; i++)
            {
                Slide cur = curCopy.Slides[i];
                if (!doNotDelete.Contains(cur))
                {
                    cur.Delete();
                }
            }
            curCopy.Save();
        }

        private Presentation MakeCopyFor(string department)
        {
            var pApp = Globals.ThisAddIn.Application;
            var pres = pApp.ActivePresentation;

            pres.SaveAs(department + ".pptx");
            Presentation temporaryPresentation = Globals.ThisAddIn.Application.Presentations.Open(department + ".pptx", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue);
            return temporaryPresentation;
        }

        private static List<int> GetSheetNumbers(string sheetRange)
        {
            List<int> sheetNumbers = new List<int>();
            if (sheetRange.Contains(","))
            {
                var sheets = sheetRange.Split(',');
                foreach (var sheet in sheets)
                {
                    sheetNumbers.Add(int.Parse(sheet));
                }
            }
            else
            {
                sheetNumbers.Add(int.Parse(sheetRange));
            }

            return sheetNumbers;
        }

    }
}
