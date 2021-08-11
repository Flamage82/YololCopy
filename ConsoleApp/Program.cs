using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using WindowsInput;
using WindowsInput.Native;
using TextCopy;

namespace YololCopy.ConsoleApp
{
    public class Program
    {
        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(IntPtr point);

        public static void Main(string[] args)
        {
            var text = ClipboardService.GetText();
            if (string.IsNullOrWhiteSpace(text))
            {
                Console.WriteLine("There is nothing on the clipboard to paste into Starbase");
                return;
            }

            var p = Process.GetProcessesByName("starbase").FirstOrDefault();
            if (p == null)
            {
                Console.WriteLine("Unable to find the Starbase process");
                return;
            }

            if (SetForegroundWindow(p.MainWindowHandle) == false)
            {
                Console.WriteLine("Unable to set focus to Starbase.");
                return;
            }

            var inputSimulator = new InputSimulator();

            foreach (var line in text.Split(new [] { "\r", "\n"}, StringSplitOptions.RemoveEmptyEntries))
            {
                inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.LCONTROL, VirtualKeyCode.VK_X);
                Thread.Sleep(50);
                ClipboardService.SetText(line);
                inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.LCONTROL, VirtualKeyCode.VK_V);
                Thread.Sleep(50);
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.DOWN);
            }
        }
    }
}
