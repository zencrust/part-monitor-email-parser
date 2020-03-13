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
            activeEandon[msg.AlertId] = msg;
            await msg.SendMqttMessage(mqttManager, logger);
        }

        private void RemoveInactive()
        {
            activeEandon
                .Where(p => p.Value.IsActive == false)
                .Select(x => x.Key)
                .Select(key => activeEandon.Remove(key));
        }

        public async Task SendMessage()
        {
            RemoveInactive();
            var activeVals = activeEandon.Select(x => x.Value).ToList();
            foreach (var item in activeVals)
            {
                item.CheckSla();
                await item.SendMqttMessage(mqttManager, logger);
            }
        }

        public void Remove(int sla)
        {
            activeEandon
                .Where(p => p.Value.SlaLevel >= sla)
                .Select(x => x.Key)
                .Select(key => activeEandon.Remove(key));
        }

        public async Task Acknowledge(EAndonMessage eAndon)
        {
            EAndonMessage oldMsg;
            if(!activeEandon.TryGetValue(eAndon.AlertId, out oldMsg))
            {
                oldMsg = eAndon;
            }
            oldMsg.Acknowledge(eAndon.InitiatedBy, eAndon.InitiateTime, eAndon.SlaLevel);
            activeEandon[eAndon.AlertId] = oldMsg;
            await oldMsg.SendMqttMessage(mqttManager, logger);
            Console.WriteLine($"{oldMsg.AlertId} : {oldMsg.AcknowledgeBy} has raised been acknowledged");
        }

        public async Task Resolve(EAndonMessage eAndon)
        {
            EAndonMessage oldMsg;
            if (!activeEandon.TryGetValue(eAndon.AlertId, out oldMsg))
            {
                oldMsg = eAndon;
            }

            oldMsg.Resolve(eAndon.InitiatedBy, eAndon.InitiateTime, eAndon.SlaLevel);
            activeEandon[eAndon.AlertId] = oldMsg;

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
