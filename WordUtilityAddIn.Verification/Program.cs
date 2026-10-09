using System;
using System.Globalization;
using System.IO;
using WordUtilityAddIn.Core.Models;
using WordUtilityAddIn.Core.Services;

namespace WordUtilityAddIn.Verification
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            if (args.Length < 2 || args.Length > 3)
            {
                Console.Error.WriteLine("用法: WordUtilityAddIn.Verification <源文档> <测试副本> [拼音字号磅数]");
                return 2;
            }

            try
            {
                string sourcePath = Path.GetFullPath(args[0]);
                string copyPath = Path.GetFullPath(args[1]);
                double points = args.Length == 3
                    ? double.Parse(args[2], CultureInfo.InvariantCulture)
                    : 9d;
                int halfPoints = checked((int)Math.Round(points * 2d, MidpointRounding.AwayFromZero));

                if (string.Equals(sourcePath, copyPath, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("测试副本路径不能与源文档相同。");
                }

                Directory.CreateDirectory(Path.GetDirectoryName(copyPath));
                File.Copy(sourcePath, copyPath, true);

                var processor = new OoxmlRubyProcessor();
                RubyStatistics source = processor.ScanPackage(sourcePath);
                RubyUpdateResult result = processor.UpdatePackage(
                    copyPath,
                    new RubyFormatOptions
                    {
                        PronunciationHalfPoints = halfPoints,
                        PronunciationFont = string.Empty,
                        AutoRaise = true
                    });

                Console.WriteLine("源文档: " + sourcePath);
                Console.WriteLine("测试副本: " + copyPath);
                Console.WriteLine("注音数量: " + result.After.RubyCount);
                Console.WriteLine("修改数量: " + result.ModifiedRubyCount);
                Console.WriteLine("拼音字号: {0} 磅 (OOXML {1})", halfPoints / 2d, result.After.DominantPronunciationHalfPoints);
                Console.WriteLine("正文基准: {0} 磅 (OOXML {1})", ToPoints(result.After.DominantBaseTextHalfPoints), Format(result.After.DominantBaseTextHalfPoints));
                Console.WriteLine("上移参数: " + Format(result.After.DominantRaiseHalfPoints));
                foreach (var part in result.After.PartRubyCounts)
                {
                    Console.WriteLine("部件注音: {0} = {1}", part.Key, part.Value);
                }
                Console.WriteLine("读音和数量校验: 通过");
                Console.WriteLine("源文档保持只读扫描，未写入。");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("验证失败: " + exception.Message);
                Console.Error.WriteLine(exception);
                return 1;
            }
        }

        private static string ToPoints(int? halfPoints)
        {
            return halfPoints.HasValue
                ? (halfPoints.Value / 2d).ToString("0.#", CultureInfo.InvariantCulture)
                : "未知";
        }

        private static string Format(int? value)
        {
            return value.HasValue ? value.Value.ToString(CultureInfo.InvariantCulture) : "未知";
        }
    }
}
