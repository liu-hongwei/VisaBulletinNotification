using System;
using System.Configuration;
using System.IO;

namespace VisaBulletinNotification
{
    static class Program
    {
        static void Main()
        {
            var watcher = new VisaBulletinWatcher();
            if (watcher.CheckNextBulletin() == true)
            {
                Logger.LogInfo("next visa bulletin is now available");
            }

            var bulletinRoot = ConfigurationManager.AppSettings.Get("VisaBulletinFilesRootPath");
            if (!Directory.Exists(bulletinRoot))
            {
                Directory.CreateDirectory(bulletinRoot);
            }

            watcher.SaveAndNotifyBulletins(bulletinRoot);
        }
    }
}
