using System;
using System.Collections.Generic;
using System.Text;

namespace DigitalVPM_DataCollector
{

    public struct APP_SETTING
    {
        public string DBPath { get; set; }
        public int DownloadTimeout { get; set; }
    }

    public class Settings
    {
        public APP_SETTING AppSettings;

        public void LoadSettings()
        {
            AppSettings.DBPath = $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\Documents\\";
            AppSettings.DownloadTimeout = 1;
        }

        public APP_SETTING GetSettings()
        {
            return AppSettings;
        }

        public void SetDownloadTimeout(int NewTime)
        {
            AppSettings.DownloadTimeout = NewTime;
        }
    }
}
