using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using DocGen.Model;

namespace DocGen.Utils
{
    static class JsonHelper
    {

        private static Settings deserialized;
        private static bool isChecked = false;

        private static string assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
        private static string assemblyFolder = System.IO.Path.GetDirectoryName(assemblyLocation);
        private static string settingsPath = System.IO.Path.Combine(assemblyFolder, "settings.txt");

        public static void SerializeJson(Settings settings)
        {
            JsonSerializer serializer = new JsonSerializer();
            serializer.NullValueHandling = NullValueHandling.Ignore;
            StreamWriter sw = null;
            try
            {
                sw = new StreamWriter(settingsPath);
                JsonWriter writer = new JsonTextWriter(sw);
                serializer.Serialize(writer, settings);
            }
            catch (Exception e)
            {
                //to-do: show message for user
                Debug.WriteLine(e.StackTrace);
            }
            finally
            {
                if (sw != null)
                {
                    sw.Close();
                }
            }

        }

        public static Settings GetDeserialized()
        {
            Settings settings = null;
            try
            {
                string json = File.ReadAllText(settingsPath);
                settings = JsonConvert.DeserializeObject<Settings>(json);
            }
            catch (FileNotFoundException e)
            {
                // if file not found, new Settings object will be created
            }
            catch (Exception e)
            {
                //to-do: show message for user
                Debug.WriteLine(e.StackTrace);
            }

            return settings;
        }

        public static Settings DeserializeJson()
        {
            if (isChecked == false)
            {
                deserialized = GetDeserialized();
                isChecked = true;
            }
            return deserialized;
        }

    }
}
