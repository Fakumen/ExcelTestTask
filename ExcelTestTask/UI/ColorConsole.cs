using System;

namespace ExcelTestTask.UI
{
    public static class ColorConsole
    {
        public static void WriteLine(object obj, System.ConsoleColor color)
        {
            var prevColor = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.WriteLine(obj);
            Console.ForegroundColor = prevColor;
        }

        public static void Write(object obj, System.ConsoleColor color)
        {
            var prevColor = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.Write(obj);
            Console.ForegroundColor = prevColor;
        }
    }
}
