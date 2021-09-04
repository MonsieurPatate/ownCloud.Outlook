using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ownCloud.Outlook
{
    /// <summary>
    /// Configuration file manager
    /// </summary>
    public class ConfigManager
    {
        /// <summary>
        /// Read config
        /// </summary>
        /// <returns></returns>
        public static Config Read()
        {
            var settings = new Config();
            var configPath = Path.Combine(Path.GetTempPath(), Constants.ConfigName);
            if (File.Exists(configPath))
            {
                using (var file = File.OpenText(configPath))
                using (var reader = new JsonTextReader(file))
                {
                    var config = ((JObject)JToken.ReadFrom(reader)).ToObject<Config>();
                    settings.Server = config.Server;
                }
            }

            return settings;
        }

        /// <summary>
        /// Save config
        /// </summary>
        /// <param name="config"></param>
        public static void Save(Config config)
        {
            if (config == null) return;
            var configPath = Path.Combine(Path.GetTempPath(), Constants.ConfigName);
            File.WriteAllText(configPath, JsonConvert.SerializeObject(config));
        }
    }
}