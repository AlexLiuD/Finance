using Quartz;
using Quartz.Impl;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace Finance.Models
{
    /// <summary>
    /// 调用京东API的Job
    /// </summary>
    public class JDJob: IJob
    {
        public void Execute(IJobExecutionContext context)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string fileName = "logServer.txt";
            string savePath = Path.Combine(path, fileName);
            if (!System.IO.File.Exists(savePath))
            {

                System.IO.FileStream fsnew = System.IO.File.Create(savePath);
                fsnew.Close();
            }
            using (System.IO.FileStream fs = System.IO.File.Open(savePath, System.IO.FileMode.Append, System.IO.FileAccess.Write))
            {
                using (System.IO.StreamWriter w = new System.IO.StreamWriter(fs))
                {
                    w.WriteLine("记录时间:{0}", DateTime.Now.ToString());
                    w.WriteLine("---------------------------------------------------------------------------------------------");
                    w.Flush();
                    w.Close();
                }
            }       
        }
    }

    public class JobScheduler
    {
        public static void Start()
        {
            //调度器
            IScheduler scheduler = StdSchedulerFactory.GetDefaultScheduler();

            //job详情
            IJobDetail job = JobBuilder.Create<JDJob>()
            .WithIdentity("job1", "group1")//定义name/group
            .Build();

            /*
            startTimeOfDay 每天开始时间
            endTimeOfDay 每天结束时间
            daysOfWeek 需要执行的星期
            interval 执行间隔
            intervalUnit 执行间隔的单位（秒，分钟，小时，天，月，年，星期）
            repeatCount 重复次数
            */

            //触发器
            ITrigger trigger = TriggerBuilder.Create()
                .WithIdentity("trigger1", "group1")//定义name/group
                .WithDailyTimeIntervalSchedule
                  (s =>
                     s.WithIntervalInMinutes(1) //每5小时执行一次
                    .OnEveryDay()
                    .StartingDailyAt(TimeOfDay.HourAndMinuteOfDay(0, 0))  //每天0:00开始
                  )
                .Build();

            //关联job和触发器  
            scheduler.ScheduleJob(job, trigger);

            //执行  
            scheduler.Start();
        }
    }
}