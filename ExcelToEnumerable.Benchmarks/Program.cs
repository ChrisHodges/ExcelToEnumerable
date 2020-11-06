using System;
using BenchmarkDotNet.Running;

namespace ExcelToEnumerable.Benchmarks.Core
{
    class Program
    {
        public static void Main(string[] args)
        {
            BenchmarkRunner.Run<Benchmarks>();
        }
    }
}