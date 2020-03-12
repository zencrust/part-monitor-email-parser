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
        List<EAndonMessage> pendingAdd = new List<EAndonMessage>();
        List<EAndonMessage> pendingRemove = new List<EAndonMessage>();

        private readonly MqttManager mqttManager;
        private readonly ILogger logger;

        public EandonManager(MqttManager mqttManager, ILogger logger)
        {
            this.mqttManager = mqttManager;
            this.logger = logger;
        }
        public async Task Initiate(EAndonMessage msg)
        {
            pendingAdd.Add(msg);
            await msg.SendMqttMessage(mqttManager, logger);
        }

        private void DoPendingAddRemove()
        {
            foreach (var item in pendingAdd)
            {
                activeEandon[item.AlertId] = item;
            }

            foreach (var item in pendingRemove)
            {
                activeEandon.Remove(item.AlertId);
            }

            activeEandon
                .Where(p => p.Value.IsActive == false)
                .Select(x => x.Key)
                .Select(key => activeEandon.Remove(key));
        }

        public async Task SendActiveMqttMessage()
        {
            DoPendingAddRemove();

            foreach (var item in activeEandon)
            {
                item.Value.CheckSla();
                await item.Value.SendMqttMessage(mqttManager, logger);
            }
        }

        public void Remove(int sla)
        {
            pendingRemove.AddRange(activeEandon
                .Where(p => p.Value.SlaLevel >= sla)
                .Select(x => x.Value));
        }

        public async Task Acknowledge(EAndonMessage eAndon)
        {
            EAndonMessage oldMsg;
            var available = activeEandon.TryGetValue(eAndon.AlertId, out oldMsg);
            oldMsg.Acknowledge(eAndon.InitiatedBy, eAndon.InitiateTime, eAndon.SlaLevel);

            if (available)
            {
                activeEandon[eAndon.AlertId] = oldMsg;
            }
            else
            {
                this.pendingAdd.Add(oldMsg);
            }

            await oldMsg.SendMqttMessage(mqttManager, logger);
            Console.WriteLine($"{oldMsg.AlertId} : {oldMsg.AcknowledgeBy} has raised been acknowledged");
        }

        public async Task Resolve(EAndonMessage eAndon)
        {
            EAndonMessage oldMsg;
            var available = activeEandon.TryGetValue(eAndon.AlertId, out oldMsg);
            oldMsg.Resolve(eAndon.InitiatedBy, eAndon.InitiateTime, eAndon.SlaLevel);

            if (available)
            {
                activeEandon[eAndon.AlertId] = oldMsg;
            }
            else
            {
                this.pendingAdd.Add(oldMsg);
            }

            await oldMsg.SendMqttMessage(mqttManager, logger);
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
