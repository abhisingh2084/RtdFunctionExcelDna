using ExcelDna.Integration.Rtd;
using static ExcelDna.Integration.Rtd.ExcelRtdServer;
using System.Collections.Generic;
using System.Timers;
using System;

public class TimeServer : ExcelRtdServer
{
    private readonly List<Topic> topics = new List<Topic>();
    private Timer timer;

    public TimeServer()
    {
        timer = new Timer();
        timer.Elapsed += TimerElapsed;
    }

    private void Start()
    {
        timer.Interval = 500; // Set the interval to 500 milliseconds
        timer.Start();
    }

    private void Stop()
    {
        timer.Stop();
    }

    private void TimerElapsed(object sender, ElapsedEventArgs e)
    {
        foreach (Topic topic in topics)
            topic.UpdateValue(GetTime());
    }

    private static string GetTime()
    {
        return DateTime.Now.ToString("HH:mm:ss.fff");
    }

    protected override void ServerTerminate()
    {
        timer.Dispose();
        timer = null;
    }

    protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
    {
        topics.Add(topic);
        Start();
        return GetTime();
    }

    protected override void DisconnectData(Topic topic)
    {
        topics.Remove(topic);
        if (topics.Count == 0)
            Stop();
    }
}
