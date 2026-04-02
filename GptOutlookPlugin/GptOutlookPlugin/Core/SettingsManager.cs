using System;
using System.IO;
using Newtonsoft.Json;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Core
{
    public class SettingsManager
    {
        private readonly string _settingsPath;
        public AppSettings Current { get; private set; }

        public SettingsManager()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var dir = Path.Combine(appData, "GptOutlookPlugin");
            Directory.CreateDirectory(dir);
            _settingsPath = Path.Combine(dir, "appsettings.json");
            Load();
        }

        public void Load()
        {
            if (File.Exists(_settingsPath))
            {
                var json = File.ReadAllText(_settingsPath);
                Current = JsonConvert.DeserializeObject<AppSettings>(json) ?? new AppSettings();
            }
            else
            {
                Current = new AppSettings();
                Save();
            }
        }

        public void Save()
        {
            var json = JsonConvert.SerializeObject(Current, Formatting.Indented);
            File.WriteAllText(_settingsPath, json);
        }
    }
}
