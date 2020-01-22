using System;
using System.Activities.Statements;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Windows.Forms;


namespace SendSheetsViaMail
{
    public partial class RibbonMenu
    {
        private void RibbonMenu_Load(object sender, RibbonUIEventArgs e)
        {

        }
        
        private void btnSetCsv_Click(object sender, RibbonControlEventArgs e)
        {
            CSVInput form = new CSVInput();
            form.Show();
        }

        private void btnSendMails_Click(object sender, RibbonControlEventArgs e)
        {
            CSVInput form = new CSVInput();
            form.Show();
           
            string subject = form.txtSubject.Text;
            string body = form.txtBody.Text;
            foreach(DataGridViewRow row in form.grdData.Rows)
            {
                if (row.IsNewRow) continue;
                string department = row.Cells["Department"].Value as string;
                string sheetRange = row.Cells["Slides"].Value as string;
               string mail = row.Cells["E-Mails"].Value as string;

                List<int> sheetNumbers = GetSheetNumbers(sheetRange);

                Presentation curCopy = MakeCopyFor(department);
                StripSheets(curCopy, sheetNumbers, department);
                PrepareMail(mail, department, subject, body);
            }
            form.Hide();

        }

        private void PrepareMail( string mail, string department, string subject, string body)
        {
           Outlook.Application oApp = new Outlook.Application();
            Outlook._MailItem oMailItem = (Outlook._MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);


            if (mail.Contains(","))
            {
                mail = mail.Replace(",", ";"); //convert to semicolon format for multiple receipients
            }

            oMailItem.To = mail;
            oMailItem.Subject = subject.Replace("$DEPARTMENT", department);
            oMailItem.Body = body.Replace("$DEPARTMENT", department);
            string path = Directory.GetCurrentDirectory();
            oMailItem.Attachments.Add(path + "/" + department + ".pptx");           
            oMailItem.Display(false);
        }

        private void StripSheets(Presentation curCopy, List<int> sheetNumbers, string department)
        {
            List<Slide> doNotDelete = new List<Slide>();
          foreach(int number in sheetNumbers)
            {
                var slide = curCopy.Slides[number];
                doNotDelete.Add(slide);
            }
            List<Slide> SlidesToDelete = new List<Slide>();
          for(int i = 1; i < curCopy.Slides.Count + 1; i++)
            {
                Slide cur = curCopy.Slides[i];
                if (!doNotDelete.Contains(cur))
                {
                    SlidesToDelete.Add(cur);
                }
            }
            SlidesToDelete.ForEach(s => s.Delete());
            curCopy.SaveAs(department + ".pptx");
            curCopy.Close();
            
            //File.Delete(department + "_raw.pptx");
        }

        private Presentation MakeCopyFor(string department)
        {
            var pApp = Globals.ThisAddIn.Application;
            var pres = pApp.ActivePresentation;

            pres.SaveAs(department + "_raw.pptx");
            var app = new Microsoft.Office.Interop.PowerPoint.Application();
            var newPresentation = app.Presentations;

            Presentation temporaryPresentation = newPresentation.Open(department + "_raw.pptx", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
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
