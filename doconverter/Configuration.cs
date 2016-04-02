using System.Configuration;

namespace doconverter
{
    static class Configuration
    {      
        public static string ReadSetting(string key)
        {
            try
            {
                var appSettings = ConfigurationManager.AppSettings;
                string result = appSettings[key] ?? "Not Found";
                return result;
            }
            catch (ConfigurationErrorsException ex)
            {
                return "Error reading app settings: " + ex.Message;
            }
        }       
    }
}

