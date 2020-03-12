using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Outlook;
using uPLibrary.Networking.M2Mqtt;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;
using NLog;
using uPLibrary.Networking.M2Mqtt.Messages;

namespace MailParser
{

    [Serializable]
    class EAndonMessage
    {
        public string AlertId { get; private set; }

        public string Alert { get; private set; }

        public string AlertType { get; private set; }

        public string Location { get; private set; }

        public string InitiatedBy { get; private set; }

        public string AcknowledgeBy { get; private set; }

        public string ResolvedBy { get; private set; }

        public string InitiateTime { get; private set; }

        public bool IsActive { get; private set; }

        public string AcknowledgeTime { get; private set; }

        public string ResolvedTime { get; private set; }

        public int SlaLevel { get; set; }

        public EAndonMessage(string alertId, string alert, string alertType, string location, string initiatedBy, string initiateTime, int slaLevel)
        {
            this.AlertId = alertId;
            this.Alert = alert;
            this.AlertType = alertType;
            this.Location = location;
            this.InitiateTime = initiateTime;
            this.InitiatedBy = initiatedBy;
            this.SlaLevel = slaLevel;
            this.IsActive = true;
        }

        public void Acknowledge(string user, string timestamp, int sla)
        {
            this.AcknowledgeBy = user;
            this.AcknowledgeTime = timestamp;
            this.IsActive = true;
            this.SlaLevel = sla;
        }

        public void Resolve(string user, string timestamp, int sla)
        {
            this.ResolvedBy = user;
            this.ResolvedTime = timestamp;
            this.IsActive = false;
            this.SlaLevel = sla;
        }

        public void ChangeSla(int sla)
        {
            if (this.SlaLevel < sla)
            {
                this.SlaLevel = sla;
            }
        }

        internal void CheckSla()
        {
            if (this.InitiateTime != "")
            {
                var initiatedTime = DateTime.Parse(this.InitiateTime);
                var calcSla = Math.Min(((int)(DateTime.Now - initiatedTime).TotalMinutes) / 30, 2);
                if (calcSla > this.SlaLevel)
                {
                    this.SlaLevel = calcSla;
                }
            }
        }
    }

    internal static class Extensions
    {
        public static string GetValue(this IList<string> inputArray, string titleString)
        {
            var result = inputArray.First(x => x.Contains(titleString))
                    .Substring(titleString.Length);
            return result.Trim();
        }

        public static string GetTimeStamp(this string timeStamp)
        {
            var timestampIndex = timeStamp.IndexOf("AM");
            if (timestampIndex == -1)
            {
                timestampIndex = timeStamp.IndexOf("PM");
            }
            return timeStamp.Substring(0, timestampIndex + 2);

        }
    }

    internal class DisposableMqttClient : MqttClient, IDisposable
    {
        public DisposableMqttClient(string brokerHostName, int brokerPort)
            : base(brokerHostName, brokerPort, false, MqttSslProtocols.None, null, null)
        {

        }
        public DisposableMqttClient(string ipaddress)
            : base(ipaddress)
        { }

        public void Dispose()
        {
            if (this.IsConnected)
            {
                this.Disconnect();
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            InitializeLogging();
            var Logger = NLog.LogManager.GetCurrentClassLogger();

            while (true)
            {
                try
                {
                    string brokerAddress = Properties.Resources.ServerAddress;
                    string clientId = Properties.Resources.ClientId;
                    Application outlookApplication = new Application();
                    NameSpace outlookNamespace = outlookApplication.GetNamespace("MAPI");
                    MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                    using (var client = new DisposableMqttClient(brokerAddress, 1883))
                    {
                        var will_topic = $"{ Properties.Resources.MqttApplicationName}/will_message/{Properties.Resources.ClientId}";
                        client.Connect(clientId, "", "", true, MqttMsgBase.QOS_LEVEL_EXACTLY_ONCE, true, will_topic, "0", true, 60);
                        Items mailItems = inboxFolder.Items;
                        var activeStations = new Dictionary<string, EAndonMessage>();
                        Logger.Info("email monitoring program started.");
                        try
                        {
                            foreach (MailItem item in mailItems)
                            {
                                Console.WriteLine(item.Body);
                                ParseEmail(client, activeStations, item.Body);
                                item.Delete();
                            }

                            bool haltMqtt = false;
                            mailItems.ItemAdd += (item) =>
                            {
                                haltMqtt = true;
                                var mailItem = item as MailItem;
                                Logger.Debug(mailItem.Body);

                                ParseEmail(client, activeStations, mailItem.Body);
                                mailItem.Delete();
                            };

                            while (true)
                            {
                                foreach (var item in activeStations)
                                {
                                    //dont run when adding email
                                    if (haltMqtt)
                                    {
                                        break;
                                    }

                                    //Application Name/Station Name/function/name
                                    item.Value.CheckSla();
                                    client.Publish(FormatMqttTopic(item.Value), GetStationSerilized(item.Value));
                                }

                                client.Publish(will_topic, Encoding.ASCII.GetBytes("1"));
                                System.Threading.Thread.Sleep(1000);
                            }
                        }
                        catch (System.Exception ex)
                        {
                            Logger.Error("parsing error");
                            Logger.Error(ex);
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Logger.Fatal(ex);
                }
            }
        }

        private static void InitializeLogging()
        {
            var config = new NLog.Config.LoggingConfiguration();

            // Targets where to log to: File and Console
            var logfile = new NLog.Targets.FileTarget("logfile") { FileName = "log.txt" };
            var logconsole = new NLog.Targets.ConsoleTarget("logconsole");

            // Rules for mapping loggers to targets            
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, logconsole);
            config.AddRule(LogLevel.Error, LogLevel.Fatal, logfile);

            // Apply config           
            NLog.LogManager.Configuration = config;
        }

        private static void ClearAllMessages(MqttClient client, Dictionary<string, EAndonMessage> activeStations)
        {
            var msg = $"{ Properties.Resources.MqttApplicationName}/{Properties.Resources.ClientId}/clear";
            client.Subscribe(new string[] { msg }, new byte[] { MqttMsgBase.QOS_LEVEL_EXACTLY_ONCE });
            client.MqttMsgPublished += (sender, e) =>
            {
            };
        }


        private static void ParseEmail(MqttClient client, Dictionary<string, EAndonMessage> activeStations, string body)
        {
            var bodyLines = body.Trim().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).ToList();
            if (!client.IsConnected)
            {
                client.Connect(Properties.Resources.ClientId);
            }
            try
            {
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

                if (status.Contains("Initiated"))
                {
                    activeStations[alertId] = stationMsg;
                    var msg = GetStationSerilized(stationMsg);
                    var res = client.Publish(FormatMqttTopic(stationMsg), msg, 2, false);
                    Console.WriteLine($"{location} : {userName} has raised eAndOn");
                }
                else if (status.Contains("Acknowledged"))
                {
                    EAndonMessage station = stationMsg;
                    activeStations.TryGetValue(alertId, out station);

                    station.Acknowledge(userName, timeStamp, slaLevel);
                    activeStations[alertId] = station;

                    var msg = GetStationSerilized(station);
                    client.Publish(FormatMqttTopic(stationMsg), msg, 2, false);
                    Console.WriteLine($"{location} : {userName} has raised been acknowledged");
                }
                else if (status.Contains("Resolved"))
                {
                    EAndonMessage station = stationMsg;
                    if (activeStations.TryGetValue(alertId, out station))
                    {
                        activeStations.Remove(alertId);
                    }

                    station.Resolve(userName, timeStamp, slaLevel);
                    var msg = GetStationSerilized(station);
                    client.Publish(FormatMqttTopic(station), msg, 2, false);
                    Console.WriteLine($"{location} : {userName} eAndOn has been closed");

                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private static byte[] GetStationSerilized(EAndonMessage message)
        {
            var str = JsonConvert.SerializeObject(message);
            return Encoding.ASCII.GetBytes(str);
        }

        private static string FormatMqttTopic(EAndonMessage station)
        {
            return $"{Properties.Resources.MqttApplicationName}/eAndon/{station.Location}/{station.Alert}";
        }

    }
}
