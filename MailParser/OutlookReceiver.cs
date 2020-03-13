using System;
using System.Linq;
using NLog;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace MailParser
{
    internal class OutlookReceiver
    {
        private readonly Application outlookApplication;
        private readonly NameSpace outlookNamespace;
        private readonly MAPIFolder inboxFolder;
        private readonly Items mailItems;
        private readonly ILogger logger;

        public OutlookReceiver(ILogger logger)
        {
            outlookApplication = new Application();
            outlookNamespace = outlookApplication.GetNamespace("MAPI");
            inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            this.logger = logger;
            mailItems = inboxFolder.Items;
        }

        public async Task ParseInbox(Func<EandonStatus, EAndonMessage, Task> action)
        {
            try
            {
                foreach (MailItem item in mailItems)
                {
                    logger.Debug(item.Body);
                    var (status, msg) = ParseEmailBody(item.Body);
                    await action(status, msg);

                    item.Delete();
                }
            }
            catch (System.Exception e)
            {
                logger.Error(e);
            }
        }

        public void RegisterForEmail(Func<EandonStatus, EAndonMessage, Task> action)
        {
            mailItems.ItemAdd += (item) =>
            {
                var mailItem = item as MailItem;
                logger.Debug(mailItem.Body);


                var (status, msg) = ParseEmailBody(mailItem.Body);
                action(status, msg).GetAwaiter().GetResult();
                mailItem.Delete();
            };
        }

        internal void Process()
        {
            outlookNamespace.SendAndReceive(false);
        }

        private static (EandonStatus, EAndonMessage) ParseEmailBody(string body)
        {
            var bodyLines = body.Trim().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).ToList();
            var alertId = bodyLines.GetValue("Alert ID:");
            var alert = bodyLines.GetValue("Alert:");
            var alertType = bodyLines.GetValue("Alert Type:");
            var location = bodyLines.GetValue("Location:");
            var status = bodyLines.GetValue("Status:");
            var slaLevel = int.Parse(bodyLines.GetValue("SLA Level:"));

            var latestHistory = bodyLines[bodyLines.Count - 3]
                .Split(new string[] { " / " }, StringSplitOptions.RemoveEmptyEntries);

            var userName = latestHistory.Last().Split(':')[0];
            var timeStamp = latestHistory.First().GetTimeStamp();

            var stationMsg = new EAndonMessage(
                alertId: alertId,
                alert: alert,
                alertType: alertType,
                location: location,
                initiatedBy: userName,
                initiateTime: timeStamp,
                slaLevel: slaLevel);

            EandonStatus enumStatus = EandonStatus.Unknown;
            if (status.Contains("Initiated"))
            {
                enumStatus = EandonStatus.Initiated;
            }
            else if (status.Contains("Acknowledged"))
            {
                enumStatus = EandonStatus.Acknowledge;
            }
            else if (status.Contains("Resolved"))
            {
                enumStatus = EandonStatus.Resolved;
            }
            return (enumStatus, stationMsg);
        }
    }
}
