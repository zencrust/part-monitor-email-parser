using Nito.Disposables;
using NLog;
using System;
using System.Globalization;
using System.Threading.Tasks;
using System.Timers;

namespace MailParser
{
    class Program
    {
        static async Task Main()
        {
            var (logger, disposable) = InitializeLogger();
            using (disposable)
            {
                try
                {
                    var mailReceiver = new OutlookReceiver(logger);
                    var taskScheduler = TaskScheduler.FromCurrentSynchronizationContext();
                    using (var mqttManager = new MqttManager(logger))
                    {
                        var eandonManager = new EandonManager(mqttManager, logger);
                        await eandonManager.Load().ConfigureAwait(false);
                        await mqttManager.Connect().ConfigureAwait(false);
                        try
                        {
                            async Task ChangeStatus(EandonStatus status, EAndonMessage msg)
                            {
                                await mqttManager.ReconnectIfNeeded().ConfigureAwait(true);
                                await eandonManager.ChangeStatus(status, msg).ConfigureAwait(false);
                            }

                            await mailReceiver.ParseInbox(ChangeStatus).ConfigureAwait(true);
                            mailReceiver.RegisterForEmail(ChangeStatus);

                            await RegisterRemoveSla(mqttManager, eandonManager, logger).ConfigureAwait(false);

                            logger.Info("email monitoring program started.");

                            using (var timer = new Timer(2000))
                            {
                                async void SendPeriodic(Object source, ElapsedEventArgs e)
                                {
                                    try
                                    {
                                        await mqttManager.ReconnectIfNeeded().ConfigureAwait(false);
                                        await eandonManager.SendMessage().ConfigureAwait(false);
                                        await mqttManager.SendOK().ConfigureAwait(false);
                                        await eandonManager.Save().ConfigureAwait(false);
                                    }
                                    catch (Exception ex)
                                    {
                                        logger.Error(ex);
                                    }
		                            finally
		                            {
		                                timer.Start();
		                            }
                                }

                                timer.AutoReset = false;
                                timer.Elapsed += new ElapsedEventHandler(SendPeriodic);
                                timer.Start();
                            }
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
                    Environment.Exit(ex.HResult);

                }
            }
        }

        private static (ILogger, IDisposable) InitializeLogger()
        {
            var config = new NLog.Config.LoggingConfiguration();

            // Targets where to log to: File and Console
            var logfile = new NLog.Targets.FileTarget("logfile") { FileName = "log.txt" };
            var logconsole = new NLog.Targets.ConsoleTarget("logconsole");

            var dispose = new AnonymousDisposable(() =>
            {
                logfile.Dispose();
                logconsole.Dispose();
            });

            // Rules for mapping loggers to targets            
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, logconsole);
            config.AddRule(LogLevel.Error, LogLevel.Fatal, logfile);

            // Apply config           
            NLog.LogManager.Configuration = config;

            return (NLog.LogManager.GetCurrentClassLogger(), dispose);
        }

        private async static Task RegisterRemoveSla(MqttManager mqttManager, EandonManager eandonManager, ILogger logger)
        {
            
            async Task remove(string topic, byte[] payload)
            {
                try
                {
                    var sla = int.Parse(System.Text.Encoding.Default.GetString(payload), CultureInfo.InvariantCulture);
                    logger.Info($"removing all SLAs for {sla}");
                    await mqttManager.ReconnectIfNeeded().ConfigureAwait(false);
                    await eandonManager.Remove(sla).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                }               
            }

            await mqttManager.Subscribe("clearSla", remove).ConfigureAwait(false);
        }
    }
}
