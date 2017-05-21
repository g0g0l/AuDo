using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace AuDo
{
    class MouseHook
    {

        private static LowLevelMouseProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;
        private static List<string> kbmlog = new List<string>();

        public static void Start()
        {
            _hookID = SetHook(_proc);
        }

        public static void Stop()
        {
            UnhookWindowsHookEx(_hookID);
            writeScript();
        }

        private static IntPtr SetHook(LowLevelMouseProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            string lastLine;
            //System.Windows.MessageBox.Show(((MouseMessages)wParam).ToString());
            if (nCode >= 0 && MouseMessages.WM_LBUTTONDOWN == (MouseMessages)wParam)
            {
                MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                lastLine = "LD " + hookStruct.pt.x + ", " + hookStruct.pt.y; // + " " + DateTime.Now.ToString("hh.mm.ss.ffff");
                kbmlog.Add(lastLine);
            }
            else if (nCode >= 0 && MouseMessages.WM_LBUTTONUP == (MouseMessages)wParam)
            {
                MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                lastLine = "LU " + hookStruct.pt.x + ", " + hookStruct.pt.y;
                //If previous line is LD with same coordinate then it's a simple left click or else it's a drag
                if (kbmlog.Count != 0)
                {
                    if (kbmlog[kbmlog.Count - 1].StartsWith("LD"))
                    {
                        if (kbmlog[kbmlog.Count - 1] == "LD " + hookStruct.pt.x + ", " + hookStruct.pt.y)
                        {
                            kbmlog.RemoveAt(kbmlog.Count - 1);
                            lastLine = "LC " + hookStruct.pt.x + ", " + hookStruct.pt.y;
                        }
                        else
                        {
                            MatchCollection matches = Regex.Matches(kbmlog[kbmlog.Count - 1], "[0-9]+");
                            List<string> coordinates = new List<string>();
                            coordinates.Clear();
                            foreach (Match march in matches)
                            {
                                coordinates.Add(march.Value);
                            }
                            kbmlog.RemoveAt(kbmlog.Count - 1);
                            lastLine = "DR " + Int32.Parse(coordinates[0]) + ", " + Int32.Parse(coordinates[1]) + ", " + hookStruct.pt.x + ", " + hookStruct.pt.y;
                        }
                    }
                }
                kbmlog.Add(lastLine);
                //If previous line is LD LU LD all with same coordinate then it's a double click
                //In this situation, after conversion to LC, if 2 LC with same coordinate, it's probably a double click
                if (kbmlog.Count > 1)
                {
                    if (kbmlog[kbmlog.Count - 1].StartsWith("LC") && (kbmlog[kbmlog.Count - 1] == kbmlog[kbmlog.Count - 2]))
                    {
                        string entryItem = kbmlog[kbmlog.Count - 1].Replace('L', 'D');
                        kbmlog.RemoveAt(kbmlog.Count - 1);
                        kbmlog.RemoveAt(kbmlog.Count - 1);
                        kbmlog.Add(entryItem);
                    }
                }
            }
            else if (nCode >= 0 && MouseMessages.WM_RBUTTONDOWN == (MouseMessages)wParam)
            {
                MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                lastLine = "RD " + hookStruct.pt.x + ", " + hookStruct.pt.y;
                kbmlog.Add(lastLine);
            }
            else if (nCode >= 0 && MouseMessages.WM_RBUTTONUP == (MouseMessages)wParam)
            {
                MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                lastLine = "RU " + hookStruct.pt.x + ", " + hookStruct.pt.y;
                //If previous line is RD with same coordinate then it's a simple right click
                if (kbmlog.Count != 0)
                {
                    if (kbmlog[kbmlog.Count - 1] == "RD " + hookStruct.pt.x + ", " + hookStruct.pt.y)
                    {
                        kbmlog.RemoveAt(kbmlog.Count - 1);
                        lastLine = "RC " + hookStruct.pt.x + ", " + hookStruct.pt.y;
                    }
                }
                kbmlog.Add(lastLine);
            }
            else if (nCode >= 0 && MouseMessages.WM_MOUSEWHEEL == (MouseMessages)wParam)
            {
                //System.Windows.MessageBox.Show("Recording scroll");
                MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                lastLine = "SC " + hookStruct.pt.x + ", " + hookStruct.pt.y + " " + GetWheelDeltaWParam((int)hookStruct.mouseData);
                kbmlog.Add(lastLine);
            }
            /*else if (nCode >= 0 && MouseMessages.WM_MOUSEMOVE == (MouseMessages)wParam)
            {
                MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                lastLine = "M " + hookStruct.pt.x + ", " + hookStruct.pt.y;
                if (kbmlog.Count != 0)
                {
                    if (kbmlog[kbmlog.Count - 1].StartsWith("LD"))
                        kbmlog.Add(lastLine);
                }
            }*/
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }
        public static void addExternalEvent(string entry)
        {
            kbmlog.Add(entry);
        }
        private static void writeScript()
        {
            kbmlog.RemoveAt(kbmlog.Count - 1);
            kbmlog.RemoveAt(kbmlog.Count - 1);
            if (!File.Exists(@"kbmlog.ads"))
            {
                using (FileStream fs = new FileStream("kbmlog.ads", FileMode.Create, FileAccess.Write))
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    foreach (string log in kbmlog)
                        sw.WriteLine(log);
                }
            }
            else
            {
                try
                {
                    using (FileStream fs = new FileStream("kbmlog.ads", FileMode.Append, FileAccess.Write))
                    using (StreamWriter sw = new StreamWriter(fs))
                    {
                        foreach (string log in kbmlog)
                            sw.WriteLine(log);
                    }
                }
                catch(System.Exception)
                {
                    System.Windows.MessageBox.Show("System busy");
                }
            }
        }

        private const int WH_MOUSE_LL = 14;
        private static int GetWheelDeltaWParam(int wparam) { return (int)(wparam >> 16); }

        private enum MouseMessages
        {
            WM_LBUTTONDOWN = 0x0201,
            WM_LBUTTONUP = 0x0202,
            WM_MOUSEMOVE = 0x0200,
            WM_MOUSEWHEEL = 0x020A,
            WM_RBUTTONDOWN = 0x0204,
            WM_RBUTTONUP = 0x0205
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook,
            LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
            IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
    }
}
