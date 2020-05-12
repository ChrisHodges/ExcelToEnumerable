using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Running;

namespace ExcelToEnumerable.Benchmarks
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var benchmarks = new Benchmarks();
            benchmarks.Setup();
            benchmarks.ExcelNPOIStorage();
            var summary = BenchmarkRunner.Run<Benchmarks>();
        }
    }
}