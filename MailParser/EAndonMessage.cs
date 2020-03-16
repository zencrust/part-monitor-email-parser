using System;
using System.Text;
using Newtonsoft.Json;
using NLog;
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

        internal async Task SendMqttMessage(MqttManager mqttManager)
        {
            await mqttManager.SendMessage(this.GetMqttTopic(), this.ToBytes());
        }

        internal async Task ForceRemove(MqttManager mqttManager)
        {
            this.IsActive = false;
            await mqttManager.SendMessage(this.GetMqttTopic(), this.ToBytes());
        }

        private string GetMqttTopic()
        {
            return $"{Properties.Settings.Default.MqttApplicationName}/eAndon/{this.Location}/{this.Alert}";
        }
    }
}
