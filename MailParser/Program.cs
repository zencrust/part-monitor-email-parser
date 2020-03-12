﻿using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Outlook;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;
using NLog;
using MQTTnet.Client;
using MQTTnet;
using MQTTnet.Client.Options;
using MQTTnet.Client.Subscribing;
using System.Threading.Tasks;

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

        internal byte[] ToBytes()
        {
            var str = JsonConvert.SerializeObject(this);
            return Encoding.ASCII.GetBytes(str);
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

    class Program
    {
        static async void Main(string[] args)
        {
            InitializeLogging();
            var logger = NLog.LogManager.GetCurrentClassLogger();

            while (true)
            {
                try
                {
                    string brokerAddress = Properties.Resources.ServerAddress;
                    string clientId = Properties.Resources.ClientId;
                    
                    Application outlookApplication = new Application();
                    NameSpace outlookNamespace = outlookApplication.GetNamespace("MAPI");
                    MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                    // Create a new MQTT client.
                    var factory = new MqttFactory();
                    var client = factory.CreateMqttClient();

                    var will_topic = $"{ Properties.Resources.MqttApplicationName}/will_message/{Properties.Resources.ClientId}";

                    var willMsgko = new MqttApplicationMessageBuilder()
                        .WithTopic(will_topic)
                        .WithRetainFlag(true)
                        .WithExactlyOnceQoS()
                        .WithPayload("0")
                        .Build();

                    var willMsgok = new MqttApplicationMessageBuilder()
                                .WithRetainFlag(false)
                                .WithTopic(will_topic)
                                .WithPayload("1")
                                .Build();


                    // Create TCP based options using the builder.
                    var options = new MqttClientOptionsBuilder()
                        .WithClientId(Properties.Resources.ClientId)
                        .WithTcpServer(brokerAddress, 1883)
                        .WithWillMessage(willMsgko)
                        .WithCleanSession()
                        .WithSessionExpiryInterval(60)
                        .Build();

                    var res = await client.ConnectAsync(options);
                    logger.Info(res);
                    using (client)
                    {
                        Items mailItems = inboxFolder.Items;
                        var activeStations = new Dictionary<string, EAndonMessage>();
                        logger.Info("email monitoring program started.");
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
                                logger.Debug(mailItem.Body);

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
                                    var mqttMsg = new MqttApplicationMessageBuilder()
                                                .WithExactlyOnceQoS()
                                                .WithRetainFlag(false)
                                                .WithTopic(FormatMqttTopic(item.Value))
                                                .WithPayload(item.Value.ToBytes())
                                                .Build();
                                    var pubRes = await client.PublishAsync(mqttMsg);
                                    logger.Info(pubRes);
                                }

                                var okWillMsgRes = await client.PublishAsync(willMsgok);
                                logger.Info(okWillMsgRes);

                                System.Threading.Thread.Sleep(1000);
                            }
                        }
                        catch (System.Exception ex)
                        {
                            logger.Error("parsing error");
                            logger.Error(ex);
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    logger.Fatal(ex);
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

        private async static Task ClearAllMessages(IMqttClient client, Dictionary<string, EAndonMessage> activeStations, ILogger logger)
        {
            var topic = $"{ Properties.Resources.MqttApplicationName}/{Properties.Resources.ClientId}/clearSla";

            var topicFilter = new TopicFilterBuilder()
                .WithTopic(topic)
                .Build();

            var topicSub = new MqttClientSubscribeOptionsBuilder()
                .WithTopicFilter(topicFilter)
                .Build();

            var res = await client.SubscribeAsync(topicSub);
            client.UseApplicationMessageReceivedHandler(e =>
           {
               if(e.ApplicationMessage.Topic == topic)
               {
                   try
                   {
                       var sla = int.Parse(System.Text.Encoding.Default.GetString(e.ApplicationMessage.Payload));

                       activeStations
                       .Where(p => p.Value.SlaLevel >= sla)
                       .Select(x => x.Key)
                       .Select(key => activeStations.Remove(key));
                    }
                   catch (Exception ex)
                   {
                       logger.Error(ex);
                   }
               }
           });
        }


        private async static Task ParseEmail(IMqttClient client, Dictionary<string, EAndonMessage> activeStations, string body, ILogger logger)
        {
            var bodyLines = body.Trim().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).ToList();
            if (!client.IsConnected)
            {
                await client.ReconnectAsync();
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
                    var msg = stationMsg.ToBytes();
                    var mqttMsg = new MqttApplicationMessageBuilder()
                        .WithExactlyOnceQoS()
                        .WithRetainFlag(false)
                        .WithTopic(FormatMqttTopic(stationMsg))
                        .WithPayload(msg)
                        .Build();
                    var res = await client.PublishAsync(mqttMsg);
                    logger.Info(res);
                    Console.WriteLine($"{location} : {userName} has raised eAndOn");
                }
                else if (status.Contains("Acknowledged"))
                {
                    EAndonMessage station = stationMsg;
                    activeStations.TryGetValue(alertId, out station);

                    station.Acknowledge(userName, timeStamp, slaLevel);
                    activeStations[alertId] = station;

                    var msg = station.ToBytes();
                    var mqttMsg = new MqttApplicationMessageBuilder()
                                    .WithExactlyOnceQoS()
                                    .WithRetainFlag(false)
                                    .WithTopic(FormatMqttTopic(station))
                                    .WithPayload(msg)
                                    .Build();
                    var res = await client.PublishAsync(mqttMsg);
                    logger.Info(res);
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
                    var msg = station.ToBytes();
                    var mqttMsg = new MqttApplicationMessageBuilder()
                                    .WithExactlyOnceQoS()
                                    .WithRetainFlag(false)
                                    .WithTopic(FormatMqttTopic(station))
                                    .WithPayload(msg)
                                    .Build();
                    var res = await client.PublishAsync(mqttMsg);
                    logger.Info(res);
                    Console.WriteLine($"{location} : {userName} eAndOn has been closed");

                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
                logger.Info(ex);
            }
        }



        private static string FormatMqttTopic(EAndonMessage station)
        {
            return $"{Properties.Resources.MqttApplicationName}/eAndon/{station.Location}/{station.Alert}";
        }

    }
}
