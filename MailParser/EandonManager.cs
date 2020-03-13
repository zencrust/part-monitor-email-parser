using System;
using System.Linq;
using System.Collections.Generic;
using NLog;
using System.Threading.Tasks;

namespace MailParser
{
    internal class EandonManager
    {
        private Dictionary<string, EAndonMessage> activeEandon = new Dictionary<string, EAndonMessage>();

        private readonly MqttManager mqttManager;
        private readonly ILogger logger;

        public EandonManager(MqttManager mqttManager, ILogger logger)
        {
            this.mqttManager = mqttManager;
            this.logger = logger;
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
                activeEandon
                    .Where(p => p.Value.IsActive == false)
                    .Select(x => x.Key)
                    .Select(key => activeEandon.Remove(key));
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
                    break;
                case EandonStatus.Acknowledge:
                    await this.Acknowledge(msg);
                    break;
                case EandonStatus.Resolved:
                    await this.Resolve(msg);
                    break;
            }
        }
    }
}
