using NLog;
using System;
using System.Threading.Tasks;
using System.Timers;

namespace MailParser
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var logger = InitializeLogger();
            try
            {
                var mailReceiver = new OutlookReceiver(logger);
                using (var mqttManager = new MqttManager(logger))
                {
                    var eandonManager = new EandonManager(mqttManager, logger);
                    try
                    {
                        async Task ChangeStatus(EandonStatus status, EAndonMessage msg)
                        {
                            await eandonManager.ChangeStatus(status, msg);
                        }

                        await mailReceiver.ParseInbox(ChangeStatus);
                        mailReceiver.RegisterForEmail(ChangeStatus);

                        await RegisterRemoveSla(mqttManager, eandonManager);

                        logger.Info("email monitoring program started.");
                        var timer = new Timer(1000);
                        async void SendPeriodic(Object source, ElapsedEventArgs e)
                        {
                            await mqttManager.ReconnectIfNeeded();
                            await eandonManager.SendActiveMqttMessage();
                            await mqttManager.SendOK();
                            //mailReceiver.Process();
                            timer.Start();
                        }

                        timer.AutoReset = false;
                        timer.Elapsed += new ElapsedEventHandler(SendPeriodic);
                        timer.Start();

                        ConsoleKey key = ConsoleKey.Clear;
                        while (true)
                        {
                            key = System.Console.ReadKey().Key;
                            if (key == ConsoleKey.Q)
                            {
                                logger.Info("Exiting Application upon user request");
                                break;
                            }
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

                var appLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
                System.Diagnostics.Process.Start(appLocation);

                // Closes the current process
                Environment.Exit(0);

            }
        }

        private static ILogger InitializeLogger()
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

            return NLog.LogManager.GetCurrentClassLogger();
        }

        private async static Task RegisterRemoveSla(MqttManager mqttManager, EandonManager eandonManager)
        {
            await mqttManager.Subscribe("clearSla", (topic, payload) =>
            {
                var sla = int.Parse(System.Text.Encoding.Default.GetString(payload));
                eandonManager.Remove(sla);
            });
        }
    }
}
