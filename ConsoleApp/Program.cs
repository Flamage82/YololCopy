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
                Thread.Sleep(50);
            }

            for (int i = 0; i < Rows; i++)
            {
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.HOME);
                Thread.Sleep(50);
                inputSimulator.Keyboard.KeyDown(VirtualKeyCode.LSHIFT);
                Thread.Sleep(50);
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.END);
                Thread.Sleep(50);
                inputSimulator.Keyboard.KeyUp(VirtualKeyCode.LSHIFT);
                Thread.Sleep(50);
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.DELETE);
                Thread.Sleep(50);
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.UP);
            }
        }

        public static void DoPaste(InputSimulator inputSimulator, string text)
        {
            for (int i = 0; i < Rows; i++)
            {
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.UP);
                Thread.Sleep(50);
            }
            inputSimulator.Keyboard.KeyPress(VirtualKeyCode.HOME);
            Thread.Sleep(50);
            
            foreach (var line in text.Split(new[] { Environment.NewLine }, StringSplitOptions.None))
            {
                inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.LCONTROL, VirtualKeyCode.VK_X);
                Thread.Sleep(50);
                ClipboardService.SetText(line);
                inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.LCONTROL, VirtualKeyCode.VK_V);
                Thread.Sleep(50);
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.DOWN);
                Thread.Sleep(50);
            }
        }

        public static void DoCopy(InputSimulator inputSimulator)
        {
            for (int i = 0; i < Rows; i++)
            {
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.UP);
                Thread.Sleep(50);
            }
            var sb = new StringBuilder();
            for(int i = 0; i < Rows; i++)
            {
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.HOME);
                Thread.Sleep(50);


                inputSimulator.Keyboard.KeyDown(VirtualKeyCode.LSHIFT);
                Thread.Sleep(50);
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.END);
                Thread.Sleep(50);
                inputSimulator.Keyboard.KeyUp(VirtualKeyCode.LSHIFT);
                Thread.Sleep(50);

                inputSimulator.Keyboard.KeyDown(VirtualKeyCode.LCONTROL);
                Thread.Sleep(50);
                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.VK_C);
                Thread.Sleep(50);
                inputSimulator.Keyboard.KeyUp(VirtualKeyCode.LCONTROL);
                Thread.Sleep(50);
                sb.AppendLine(ClipboardService.GetText());

                inputSimulator.Keyboard.KeyPress(VirtualKeyCode.DOWN);
                Thread.Sleep(50);
            }

            ClipboardService.SetText(sb.ToString());
        }
    }
}
