using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace AuDo
{
    class KBHandler
    {
        public static void Press(string key)
        {
            switch(key)
            {
                case "Space":
                    key = " ";
                    break;
                case "Return":
                    key = "{ENTER}";
                    break;
                case "Back":
                    key = "{BACKSPACE}";
                    break;
                case "OemMinus":
                    key = "-";
                    break;
                case "Oemplus":
                    key = "=";
                    break;
                case "Oemcomma":
                    key = ",";
                    break;
                case "OemPeriod":
                    key = ".";
                    break;
                case "OemQuestion":
                    key = "/";
                    break;
                case "LShiftKey":
                    //key = "+";//This should be it, but it is not working
                    key = "";
                    break;
                case "RShiftKey":
                    //key = "+";//Read previous comment
                    key = "";
                    break;
                case "RControlKey":
                    key = "^";
                    break;
                case "LControlKey":
                    key = "^";
                    break;
                default:
                    break;
            }
            string txt = Regex.Replace(key, "[+^%~()]", "{$0}");
            txt = txt.ToLower();
            SendKeys.SendWait(txt);
        }
    }
}
