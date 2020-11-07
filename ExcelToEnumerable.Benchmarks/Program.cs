using BenchmarkDotNet.Running;

namespace ExcelToEnumerable.Benchmarks.Core
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            BenchmarkRunner.Run<Benchmarks>();
        }
    }
}