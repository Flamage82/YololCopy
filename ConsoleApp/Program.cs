using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using WindowsInput;
using WindowsInput.Native;
using TextCopy;
using System.Text;

namespace YololCopy.ConsoleApp
{
    public class Program
    {
        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(IntPtr point);
        private static readonly int Rows = 20;
        private enum Invoke { Bad, Copy, Paste, Clear }
        private static readonly int SleepTime = 50;

        public static void Main(string[] args)
        {
            var invoke = CheckArgs(args);
            if (invoke == Invoke.Bad) return;

            var text = ClipboardService.GetText();
            if (invoke == Invoke.Paste && string.IsNullOrWhiteSpace(text))
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

            switch (invoke)
            {
                case Invoke.Copy:
                    DoCopy(inputSimulator);
                    break;
                case Invoke.Paste:
                    DoPaste(inputSimulator, text);
                    break;
                case Invoke.Clear:
                    DoClear(inputSimulator);
                    break;
            }

            Console.WriteLine("Done\a");
        }

        private static Invoke CheckArgs(string [] args)
        {
            if (args.Length != 1 || !new string[] { "--copy", "--paste", "--clear" }.Contains(args[0]))
            {
                Console.WriteLine(
                    string.Format(
                        "USAGE: {0} [--copy|--paste|--clear]",
                        System.AppDomain.CurrentDomain.FriendlyName));
                return Invoke.Bad;
            }
            switch (args[0])
            {
                case "--copy":
                    return Invoke.Copy;
                case "--paste":
                    return Invoke.Paste;
                case "--clear":
                    return Invoke.Clear;
                default:
                    throw new ArgumentException(string.Format("Unrecognized argument {0}", args[0]));
            }
        }

        public static void DoClear(InputSimulator inputSimulator)
        {
            for(int i = 0; i < Rows; i++)
            {
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.DOWN);
                Thread.Sleep(SleepTime);
            }

            for (int i = 0; i < Rows; i++)
            {
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.HOME);
                Thread.Sleep(SleepTime);
                inputSimulator.Keyboard.KeyDown(VirtualKeyCode.LSHIFT);
                Thread.Sleep(SleepTime);
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.END);
                Thread.Sleep(SleepTime);
                inputSimulator.Keyboard.KeyUp(VirtualKeyCode.LSHIFT);
                Thread.Sleep(SleepTime);
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.DELETE);
                Thread.Sleep(SleepTime);
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.UP);
            }
        }

        public static void DoPaste(InputSimulator inputSimulator, string text)
        {
            for (int i = 0; i < Rows; i++)
            {
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.UP);
                Thread.Sleep(SleepTime);
            }
            inputSimulator.Keyboard.KeyPress(VirtualKeyCode.HOME);
            Thread.Sleep(SleepTime);
            
            foreach (var line in text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
            {
                inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.LCONTROL, VirtualKeyCode.VK_X);
                Thread.Sleep(SleepTime);
                ClipboardService.SetText(line);
                inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.LCONTROL, VirtualKeyCode.VK_V);
                Thread.Sleep(SleepTime);
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.DOWN);
                Thread.Sleep(SleepTime);
            }
        }

        public static void DoCopy(InputSimulator inputSimulator)
        {
            for (int i = 0; i < Rows; i++)
            {
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.UP);
                Thread.Sleep(SleepTime);
            }
            var sb = new StringBuilder();
            for(int i = 0; i < Rows; i++)
            {
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.HOME);
                Thread.Sleep(SleepTime);


                inputSimulator.Keyboard.KeyDown(VirtualKeyCode.LSHIFT);
                Thread.Sleep(SleepTime);
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.END);
                Thread.Sleep(SleepTime);
                inputSimulator.Keyboard.KeyUp(VirtualKeyCode.LSHIFT);
                Thread.Sleep(SleepTime);

                inputSimulator.Keyboard.KeyDown(VirtualKeyCode.LCONTROL);
                Thread.Sleep(SleepTime);
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.VK_C);
                Thread.Sleep(SleepTime);
                inputSimulator.Keyboard.KeyUp(VirtualKeyCode.LCONTROL);
                Thread.Sleep(SleepTime);
                sb.AppendLine(ClipboardService.GetText());

                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.DOWN);
                Thread.Sleep(SleepTime);
            }

            ClipboardService.SetText(sb.ToString());
        }
    }
}
