using System;
using System.Linq;
using System.Collections.Generic;
using NLog;
using System.Threading.Tasks;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using Newtonsoft.Json;
using System.Text;

namespace MailParser
{
    internal class EandonManager
    {
        private Dictionary<string, EAndonMessage> activeEandon = new Dictionary<string, EAndonMessage>();

        private readonly MqttManager mqttManager;
        private readonly ILogger logger;
        private bool dirty = false;

        public EandonManager(MqttManager mqttManager, ILogger logger)
        {
            this.mqttManager = mqttManager;
            this.logger = logger;
        }

        public Task Load()
        {
            return Task.Run(() =>
            {
                try
                {
                    if (File.Exists(Properties.Settings.Default.CurrentEands))
                    {
                        using (var stream = File.OpenText(Properties.Settings.Default.CurrentEands))
                        {
                            JsonSerializer serializer = new JsonSerializer();
                            if (serializer.Deserialize(stream, typeof(EAndonMessage[])) is EAndonMessage[] data)
                            {
                                foreach (var item in data)
                                {
                                    lock (this.activeEandon)
                                    {
                                        this.activeEandon[item.AlertId] = item;
                                    }
                                }
                            }
                        }
                    }
                }
                catch(Exception ex)
                {
                    logger.Error(ex);
                }
            });
        }

        public Task Save()
        {
            return Task.Run(() =>
            {
                if (this.dirty)
                {
                    File.Delete(Properties.Settings.Default.CurrentEands);
                    using (var stream = File.CreateText(Properties.Settings.Default.CurrentEands))
                    {
                        JsonSerializer serializer = new JsonSerializer();
                        serializer.Serialize(stream, this.activeEandon.Values.ToArray());
                    }
                }

                this.dirty = false;
            });

        }

        public async Task Initiate(EAndonMessage msg)
        {
            lock (activeEandon)
            {
                activeEandon[msg.AlertId] = msg;
            }
            await msg.SendMqttMessage(mqttManager);
        }

        private void RemoveInactive()
        {
            lock (activeEandon)
            {
                var anyModified = activeEandon
                    .Where(p => p.Value.IsActive == false)
                    .Select(x => x.Key)
                    .Select(key => activeEandon.Remove(key))
                    .Any();

                if(anyModified)
                {
                    this.dirty = true;
                }
            }

        }

        public async Task SendMessage()
        {
            RemoveInactive();
            List<EAndonMessage> activeVals;
            lock (activeEandon)
            {
                activeVals = activeEandon.Select(x => x.Value).ToList();
            }

            foreach (var item in activeVals)
            {
                item.CheckSla();
                await item.SendMqttMessage(mqttManager);
            }
        }

        public async Task Remove(int sla)
        {
            List<KeyValuePair<string, EAndonMessage>> eAndonMessages;
            lock (activeEandon)
            {
                eAndonMessages = activeEandon
                    .Where(p => p.Value.SlaLevel >= sla)
                    .Select(x => x).ToList();
                foreach (var item in eAndonMessages)
                {
                    activeEandon.Remove(item.Key);
                }
            }

            foreach (var keyValuePair in eAndonMessages)
            {
                await keyValuePair.Value.ForceRemove(mqttManager);
            }

            if (eAndonMessages.Any())
            {
                this.dirty = true;
            }

        }

        public async Task Acknowledge(EAndonMessage eAndon)
        {
            EAndonMessage oldMsg;
            lock (activeEandon)
            {
                if (!activeEandon.TryGetValue(eAndon.AlertId, out oldMsg))
                {
                    oldMsg = eAndon;
                }
                oldMsg.Acknowledge(eAndon.InitiatedBy, eAndon.InitiateTime, eAndon.SlaLevel);
                activeEandon[eAndon.AlertId] = oldMsg;
            }
            await oldMsg.SendMqttMessage(mqttManager);
            Console.WriteLine($"{oldMsg.AlertId} : {oldMsg.AcknowledgeBy} has raised been acknowledged");
        }

        public async Task Resolve(EAndonMessage eAndon)
        {
            EAndonMessage oldMsg;
            lock (activeEandon)
            {               
                if (!activeEandon.TryGetValue(eAndon.AlertId, out oldMsg))
                {
                    oldMsg = eAndon;
                }

                oldMsg.Resolve(eAndon.InitiatedBy, eAndon.InitiateTime, eAndon.SlaLevel);
                activeEandon[eAndon.AlertId] = oldMsg;
            }

            await oldMsg.SendMqttMessage(mqttManager);
            Console.WriteLine($"{oldMsg.AlertId} : {oldMsg.AcknowledgeBy} has been resolved");
        }

        public async Task ChangeStatus(EandonStatus status, EAndonMessage msg)
        {
            switch (status)
            {
                case EandonStatus.Initiated:
                    await this.Initiate(msg);
                    this.dirty = true;
                    break;
                case EandonStatus.Acknowledge:
                    await this.Acknowledge(msg);
                    this.dirty = true;
                    break;
                case EandonStatus.Resolved:
                    await this.Resolve(msg);
                    this.dirty = true;
                    break;
            }
        }
    }
}
