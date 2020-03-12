using NLog;
using MQTTnet.Client;
using MQTTnet;
using MQTTnet.Client.Options;
using System.Threading.Tasks;
using MQTTnet.Client.Subscribing;
using System;

namespace MailParser
{
    internal class MqttManager : IDisposable
    {
        private readonly ILogger logger;
        private readonly IMqttClientOptions options;
        private readonly MqttApplicationMessage willMsgok;
        private readonly IMqttClient client;

        public MqttManager(ILogger logger)
        {
            string brokerAddress = Properties.Resources.ServerAddress;
            string clientId = Properties.Resources.ClientId;

            var factory = new MqttFactory();
            client = factory.CreateMqttClient();

            var will_topic = $"{ Properties.Resources.MqttApplicationName}/will_message/{Properties.Resources.ClientId}";

            var willMsgko = new MqttApplicationMessageBuilder()
                .WithTopic(will_topic)
                .WithRetainFlag(true)
                .WithExactlyOnceQoS()
                .WithPayload("0")
                .Build();

            // Create a new MQTT client.
            willMsgok = new MqttApplicationMessageBuilder()
                        .WithRetainFlag(false)
                        .WithTopic(will_topic)
                        .WithPayload("1")
                        .Build();


            // Create TCP based options using the builder.
            options = new MqttClientOptionsBuilder()
                .WithClientId(clientId)
                .WithTcpServer(brokerAddress, 1883)
                .WithWillMessage(willMsgko)
                .WithCleanSession()
                .WithSessionExpiryInterval(60)
                .Build();

            this.logger = logger;
        }

        internal async Task Connect()
        {
            var res = await client.ConnectAsync(options);
            logger.Info(res);
        }

        internal async Task ReconnectIfNeeded()
        {
            if (!client.IsConnected)
            {
                await client.ReconnectAsync();
            }
        }

        internal async Task SendMessage(string topic, byte[] message)
        {
            var mqttMsg = new MqttApplicationMessageBuilder()
                            .WithExactlyOnceQoS()
                            .WithRetainFlag(false)
                            .WithTopic(topic)
                            .WithPayload(message)
                            .Build();
                                var pubRes = await client.PublishAsync(mqttMsg);
            logger.Info(pubRes);

        }

        internal async Task Subscribe(string endTopic, Action<string, byte[]> action)
        {
            var topic = $"{ Properties.Resources.MqttApplicationName}/{Properties.Resources.ClientId}/{endTopic}";

            var topicFilter = new TopicFilterBuilder()
                .WithTopic(topic)
                .Build();

            var topicSub = new MqttClientSubscribeOptionsBuilder()
                .WithTopicFilter(topicFilter)
                .Build();

            var res = await client.SubscribeAsync(topicSub);
            client.UseApplicationMessageReceivedHandler(e =>
            {
                if (e.ApplicationMessage.Topic == topic)
                {
                    try
                    {
                        action(topic, e.ApplicationMessage.Payload);
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex);
                    }
                }
            });
        }

        internal async Task SendOK()
        {
            var okWillMsgRes = await client.PublishAsync(willMsgok);
            logger.Info(okWillMsgRes);
        }

        public void Dispose()
        {
            client.Dispose();
        }
    }
}
