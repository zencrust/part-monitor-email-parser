using NLog;
using MQTTnet.Client;
using MQTTnet;
using MQTTnet.Client.Options;
using System.Threading.Tasks;
using MQTTnet.Client.Subscribing;
using System;
using MQTTnet.Client.Publishing;
using MQTTnet.Client.Connecting;
using System.Threading;

namespace MailParser
{
    internal class MqttManager : IDisposable
    {
        private readonly ILogger logger;
        private readonly IMqttClientOptions options;
        private readonly MqttApplicationMessage willMsgok;
        private readonly IMqttClient client;
        private readonly SemaphoreSlim reconnectSync = new SemaphoreSlim(1);

        public MqttManager(ILogger logger)
        {
            string brokerAddress = Properties.Settings.Default.ServerAddress;
            string clientId = Properties.Settings.Default.ClientId;

            var factory = new MqttFactory();
            client = factory.CreateMqttClient();

            var will_topic = $"{ Properties.Settings.Default.MqttApplicationName}/will_message/{Properties.Settings.Default.ClientId}";

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
                .WithKeepAlivePeriod(TimeSpan.FromSeconds(10))
                .Build();

            this.logger = logger;
        }

        internal async Task Connect()
        {
            var res = await this.client.ConnectAsync(options).ConfigureAwait(false);
            if (res.ResultCode != MqttClientConnectResultCode.Success)
            {
                logger.Warn(res.ResultCode.ToString() + " " + res.ReasonString);
            }
        }

        internal async Task ReconnectIfNeeded()
        {
            await reconnectSync.WaitAsync().ConfigureAwait(false);

            while (!client.IsConnected)
            {
                try
                {
                    logger.Warn("mqtt disconnected! trying to reconnect");
                    await client.ReconnectAsync().ConfigureAwait(false);
                }
                catch (MQTTnet.Exceptions.MqttCommunicationException ex)
                {
                    logger.Error(ex);
                    await Task.Delay(2000).ConfigureAwait(false);
                }
            }
            reconnectSync.Release();
        }

        internal async Task SendMessage(string topic, byte[] message)
        {
            var mqttMsg = new MqttApplicationMessageBuilder()
                            .WithExactlyOnceQoS()
                            .WithRetainFlag(false)
                            .WithTopic(topic)
                            .WithPayload(message)
                            .Build();
            var res = await this.client.PublishAsync(mqttMsg).ConfigureAwait(false);
            if(res.ReasonCode != MqttClientPublishReasonCode.Success)
            {
                logger.Warn(res.ReasonCode.ToString() + " " + res.ReasonString);
            }
        }

        internal async Task Subscribe(string endTopic, Func<string, byte[], Task> action)
        {
            var topic = $"{ Properties.Settings.Default.MqttApplicationName}/{Properties.Settings.Default.ClientId}/{endTopic}";
            var topicFilter = new TopicFilterBuilder()
                .WithTopic(topic)
                .WithExactlyOnceQoS()
                .Build();

            var topicSub = new MqttClientSubscribeOptionsBuilder()
                .WithTopicFilter(topicFilter)
                .Build();

            var res = await this.client.SubscribeAsync(topicSub).ConfigureAwait(false);
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
            var res = await this.client.PublishAsync(willMsgok).ConfigureAwait(false);
            if(res.ReasonCode != MqttClientPublishReasonCode.Success)
            {
                logger.Warn(res.ReasonCode.ToString() + " " + res.ReasonString);

            }
        }

        public void Dispose()
        {
            client.Dispose();
        }
    }
}
