using System;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(typeof(string).Assembly.FullName);
            Console.WriteLine(typeof(int).Assembly.FullName);
            Console.WriteLine(typeof(bool).Assembly.FullName);
            Console.WriteLine(typeof(Func<>).Assembly.FullName);
            Console.WriteLine("Hello World!");
        }
    }
}
