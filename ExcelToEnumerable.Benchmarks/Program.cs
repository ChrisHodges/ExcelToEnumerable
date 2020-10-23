using System;
using BenchmarkDotNet.Running;

namespace ExcelToEnumerable.Benchmarks.Core
{
    class Program
    {
        public static void Main(string[] args)
        {
            var benchmarks = new Benchmarks();
            benchmarks.Setup();
            var summary = BenchmarkRunner.Run<Benchmarks>();
        }
    }
}