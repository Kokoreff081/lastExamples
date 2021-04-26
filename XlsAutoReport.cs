using System;
using System.ServiceProcess;
using Quartz;
using Quartz.Impl;
using System.Collections.Specialized;
using System.Configuration;
using System.ServiceModel;
using WcfQueryXlsx;

namespace XlsAutoReportWindowsService
{
    public partial class XlsService : ServiceBase
    {
        private static IScheduler scheduler;
       // private static TCPServer server = null;
        internal static ServiceHost hostService = null;
        public XlsService()
        {
            InitializeComponent();
        }
        public static void WorkDebug()
        {
            Init();
            DoWork();
            //        InitServer();
           
            System.Threading.Thread.Sleep(Int32.MaxValue);
        }
        protected override void OnStart(string[] args)
        {
            Init();
            DoWork();
            try
            {
                hostService = new ServiceHost(typeof(Service1));
                Service1 mgmtService = new Service1();
                hostService.Open();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            //        InitServer();
        }

        protected override void OnStop()
        {
            //        server.StopServer();
            //         server = null;
            hostService.Close();
            SLogger.Log.Info("Остановка службы");
        }
        static void Init()
        {
            NameValueCollection props = new NameValueCollection
            {
                { "quartz.serializer.type", "binary" },
                { "quartz.scheduler.instanceName", "MyScheduler" }
            };
            StdSchedulerFactory factory = new StdSchedulerFactory(props);
            scheduler = factory.GetScheduler().ConfigureAwait(false).GetAwaiter().GetResult();
        }
        static void DoWork()
        {
            SLogger.Log.Info("Старт службы");
            scheduler.Start().ConfigureAwait(false).GetAwaiter().GetResult();

            try
            {
                ScheduleJobRequest(scheduler);
                scheduler.Start();
            }
            catch (Exception ex)
            {
                SLogger.Log.Error($"Произошла ошибка при старте службы. Подробнее: {ex.Message}");
            }
        }

        static bool ScheduleJobRequest(IScheduler scheduler)
        {
#if (DEBUG)
            String shedule = "DebugSchedule";
            String shedule2 = "DebugSchedule2";
#endif

#if (!DEBUG)

            String shedule = "TrackerDailySchedule";
            String shedule2 = "TrackerWeeklySchedule";
#endif
            SLogger.Log.Info("Запуск шедулера");
            string errors = String.Empty;
            try
            {
                var jobDetail = new JobDetailImpl("job1", "group1", typeof(TrackerDailyReport));
                var jobDetai2 = new JobDetailImpl("job2", "group2", typeof(TrackerWeeklyReport));

                ITrigger trigger = TriggerBuilder.Create()                  // создаем триггер
                    .WithIdentity("job1", "group1")     // идентифицируем триггер с именем и группой

                    .StartNow()
                    .WithCronSchedule(ConfigurationManager.AppSettings[shedule], x => x.WithMisfireHandlingInstructionDoNothing())     // расписание
                    .Build();                                               // создаем триггер

                ITrigger trigger2 = TriggerBuilder.Create()                  // создаем триггер
                    .WithIdentity("job2", "group2")     // идентифицируем триггер с именем и группой

                    .StartNow()
                    .WithCronSchedule(ConfigurationManager.AppSettings[shedule2], x => x.WithMisfireHandlingInstructionDoNothing())     // расписание
                    .Build();                                               // создаем триггер    

                scheduler.ScheduleJob(jobDetail, trigger); // начинаем выполнение работы ежедневная рассылка понедельник-пятница макеты за вчера
                scheduler.ScheduleJob(jobDetai2, trigger2);// начинаем выполнение работы еженедельная рассылка макеты и зп за прошлую неделю

                return true;
            }
            catch (Exception ex)
            {
                SLogger.Log.Error($"Ошибка запуска задания: {ex.Message}");
                return false;
            }
        }
        static void InitServer()
        {
            SLogger.Log.Info("Запуск службы сервера");
  //          server = new TCPServer();
   //        server.StartServer();
        }
    }
}
