using DocGen.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGen.Model
{
    class SettingsFactory
    {
        public Settings GetSettings()
        {
            Settings deserialized = JsonHelper.DeserializeJson();
            if (deserialized == null)
            {
                return Settings.Instance();
            }
            else
            {
                return deserialized;
            }

        }
    }
}
