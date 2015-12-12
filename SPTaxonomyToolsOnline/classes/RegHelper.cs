using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BandR
{
    public class RegistryHelper
    {

        const string REG_KEY_APP_NAME = "SPTAXONOMYTOOLSONLINE";
        const string REG_KEY_VALUE = "ALLDATA";

        /// <summary>
        /// </summary>
        public static bool GetRegStuff(out string alldata, out string msg)
        {
            alldata = "";
            msg = "";

            try
            {
                var key = Registry.CurrentUser.OpenSubKey("Software");
                var appKey = key.OpenSubKey(REG_KEY_APP_NAME);

                if (appKey != null)
                {
                    alldata = GenUtil.SafeTrim(appKey.GetValue(REG_KEY_VALUE));
                }

            }
            catch (Exception ex)
            {
                msg = "Registry read error: " + ex.Message;
            }

            return msg == "";
        }

        /// <summary>
        /// </summary>
        public static bool SaveRegStuff(string alldata, out string msg)
        {
            msg = "";

            try
            {
                var key = Registry.CurrentUser.OpenSubKey("Software", true);
                var appKey = key.OpenSubKey(REG_KEY_APP_NAME, true);

                if (appKey == null)
                {
                    appKey = key.CreateSubKey(REG_KEY_APP_NAME);
                }

                appKey.SetValue(REG_KEY_VALUE, GenUtil.SafeTrim(alldata));

            }
            catch (Exception ex)
            {
                msg = "Registry save error: " + ex.Message;
            }

            return msg == "";
        }

    }
}
