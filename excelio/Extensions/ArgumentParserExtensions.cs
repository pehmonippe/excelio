namespace excelio.Extensions
{
    using System.Collections.Generic;
    using System.IO;
    using JetBrains.Annotations;

    internal static class ArgumentParserExtensions
    {
        public static string GetFullExcelFilePath ([NotNull] this IDictionary<string, object> parameters)
        {
            var fullFilePath = (string) parameters[ArgumentParser.InputFile];
            return fullFilePath;
        }

        public static string GetExcelFileName ([NotNull] this IDictionary<string, object> parameters)
        {
            var inputFile = (string) parameters[ArgumentParser.InputFile];
            var excelFileName = Path.GetFileNameWithoutExtension(inputFile);

            return excelFileName;
        }

        public static string GetOutputPath ([NotNull] this IDictionary<string, object> parameters)
        {
            var outputFolder = (string) parameters[ArgumentParser.OutputFolder];
            return outputFolder;
        }

        public static bool ShouldRandomize ([NotNull] this IDictionary<string, object> parameters)
        {
            int result;

            var shouldRandomize = int.TryParse((string)parameters[ArgumentParser.Randomize], out result)
                && 1 == result;

            return shouldRandomize;
        }
    }
}