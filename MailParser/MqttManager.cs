using NLog;
using MQTTnet.Client;
using MQTTnet;
using MQTTnet.Client.Options;
using System.Threading.Tasks;
using MQTTnet.Client.Subscribing;
using System;
using MQTTnet.Client.Publishing;
using MQTTnet.Client.Connecting;

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
            var res = await client.ConnectAsync(options);
            if (res.ResultCode != MqttClientConnectResultCode.Success)
            {
                logger.Warn(res.ResultCode.ToString() + " " + res.ReasonString);
            }
        }

        internal async Task ReconnectIfNeeded()
        {
            try
            {
                if (!client.IsConnected)
                {
                    await client.ReconnectAsync();
                }
            }
            catch (MQTTnet.Exceptions.MqttCommunicationException ex)
            {
                logger.Error(ex);
                //var appLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
                //System.Diagnostics.Process.Start(appLocation);

                // Closes the current process
                Environment.Exit(ex.HResult);
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
            var res = await client.PublishAsync(mqttMsg);
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
            var res = await client.PublishAsync(willMsgok);
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
