using System;
using System.Collections;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;


namespace Automation.TestScripts
{
    public class Notifications : TestCaseUtil
    {
        #region Constructeurs
        public Notifications()
        {
        }
        public Notifications(string reportFileDirectory, string embeddedMailContents)
        {
            this.ReportFileDirectory = reportFileDirectory;
            this.EmbeddedMailContents = embeddedMailContents;
        }
        #endregion

        string _sReportFileDirectory;
        ///
        /// Report File Directory
        ///
        public string ReportFileDirectory
        {
            get { return _sReportFileDirectory; }
            set { _sReportFileDirectory = value; }
        }

        string _sEmbeddedMailContents;
        ///
        /// Report File Directory
        ///
        public string EmbeddedMailContents
        {
            get { return _sEmbeddedMailContents; }
            set { _sEmbeddedMailContents = value; }
        }


        #region Function to send notifications to the user
        /// <summary>
        /// Function to send notifications to the user
        /// </summary>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>02-Sept-2011</Date>
        public void SendNotification()
        {
            try
            {
                string emailHTMLStart = "<HTML><BODY>";
                string emailHTMLEnd = "</BODY></HTML>";
                string emailSummery = "<strong><font size=\"3\" face=\"Calibri\"> ComputeNext Test Automation Summery </font></strong><br /><br />";

                string attachmentFilePath = this.ReportFileDirectory + ".zip";

                string embeddedImagePath = this.ReportFileDirectory + "\\HtmlFiles\\excel_chart_export.bmp";

                int portNo = 587;
                // Loads config data and creates a Singleton object of Configuration and loads data into generic test case variables
                this.GetConfigData();

                emailSubject = emailSubject + " " + DateTime.Now;

                // Create a message and set up the recipients.
                emailTo = emailTo.Replace(";", ",");

                // prepare e-mail body text with server name
                emailMessage = "<font size=\"3\" face=\"Calibri\" color=black>" + emailMessage + " " + server + " vm.</font>";

                MailMessage message = new MailMessage(emailFrom, emailTo, emailSubject, emailMessage);

                // Code to attach zip files to the e-mail.
                // For now the zip file is skipped because of zip file size issue (Do not delete this code)
                    //// See if this file exists in the directory 
                    //if (File.Exists(attachmentFilePath))
                    //{
                    //    // Create  the file attachment for this e-mail message.
                    //    Attachment data = new Attachment(attachmentFilePath, MediaTypeNames.Application.Octet);
                    //    // Add the file attachment to this e-mail message.
                    //    message.Attachments.Add(data);
                    //}

                //Send the message.
                SmtpClient client = new SmtpClient(emailSMTPServer, portNo);

                // Add credentials if the SMTP server requires them.
                client.Credentials = CredentialCache.DefaultNetworkCredentials;
                client.EnableSsl = false;
                client.Timeout = 200000;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new System.Net.NetworkCredential(emailUserName, emailPassword);

                //Creating AlternateView for Plain Text Mail
                AlternateView normalView = AlternateView.CreateAlternateViewFromString(emailMessage, null, "text/plain");

                //Creating AlternateView for HTML Mail
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(emailHTMLStart + emailMessage + "<br /><br />" + emailSummery + this.EmbeddedMailContents + "<br /><br />" + " <img src=cid:me>" + emailHTMLEnd, null, "text/html");
                //AlternateView htmlView = AlternateView.CreateAlternateViewFromString(this.EmbeddedMailContents, null, "text/html");

                // See if this file exists in the directory 
                if (File.Exists(embeddedImagePath))
                {
                    //Creating LinkedSource for embedded picture
                    LinkedResource myPhoto = new LinkedResource(embeddedImagePath);
                    myPhoto.ContentId = "me";

                    //Adding LinkedSource to AlternateView
                    htmlView.LinkedResources.Add(myPhoto);
                }

                //Adding AlternateViews to MailMessage
                message.AlternateViews.Add(normalView);
                message.AlternateViews.Add(htmlView);

                message.BodyEncoding = System.Text.UTF8Encoding.UTF8;
                message.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;

                try
                {
                    client.Send(message);
                }
                catch (Exception exception)
                {
                    Console.WriteLine("Exception caught in Send Notification(): Not able to send message. ", exception.ToString());
                }

                message.Dispose();
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception caught in Notification(): " + exception.ToString());
            }
        }
        #endregion
    }
}


