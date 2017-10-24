namespace excelio
{
    using System;
    using System.Collections.Generic;
    using JetBrains.Annotations;

    internal class ArgumentParser
    {
        public const string OutputFolder = "output-folder";
        public const string InputFile = "input-file";
        public const string FileFormat = "file-format";
        public const string Randomize = "randomize";

        private readonly IDictionary<string, object> parameters =
            new Dictionary<string, object>
            {
                { InputFile, string.Empty },
                { OutputFolder, @".\exports" },
                { FileFormat, 0 },
                { Randomize, 0 }
            };

        /// <summary>
        /// Parses the specified arguments.
        /// </summary>
        /// <param name="args">The arguments.</param>
        /// <returns></returns>
        [NotNull]
        public IDictionary<string, object> Parse ([CanBeNull] string[] args)
        {
            if (null == args)
                return parameters;

            foreach (var arg in args)
            {
                var elements = arg.Split(new [] { "--", "=" }, StringSplitOptions.RemoveEmptyEntries);

                if (elements.Length < 2)
                    continue;

                if (!parameters.ContainsKey(elements[0]))
                    continue;

                parameters[elements[0]] = elements[1];
            }

            return parameters;
        }
    }
}