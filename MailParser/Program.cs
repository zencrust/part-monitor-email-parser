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
                    await mqttManager.Connect();
                    try
                    {
                        async Task ChangeStatus(EandonStatus status, EAndonMessage msg)
                        {
                            await mqttManager.ReconnectIfNeeded();
                            await eandonManager.ChangeStatus(status, msg);
                        }

                        await mailReceiver.ParseInbox(ChangeStatus);
                        mailReceiver.RegisterForEmail(ChangeStatus);


                        await RegisterRemoveSla(mqttManager, eandonManager, logger);

                        logger.Info("email monitoring program started.");
                        var timer = new Timer(2000);
                        async void SendPeriodic(Object source, ElapsedEventArgs e)
                        {
                            try
                            {
                                await mqttManager.ReconnectIfNeeded();
                                await eandonManager.SendMessage();
                                await mqttManager.SendOK();
                                timer.Start();
                            }
                            catch(Exception ex)
                            {
                                logger.Error(ex);
                            }
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

        private async static Task RegisterRemoveSla(MqttManager mqttManager, EandonManager eandonManager, ILogger logger)
        {
            
            async Task remove(string topic, byte[] payload)
            {
                try
                {
                    var sla = int.Parse(System.Text.Encoding.Default.GetString(payload));
                    logger.Info($"removing all SLAs for {sla}");
                    await mqttManager.ReconnectIfNeeded();
                    await eandonManager.Remove(sla);
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                }               
            }

            await mqttManager.Subscribe("clearSla", remove);
        }
    }
}
