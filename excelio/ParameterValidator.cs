namespace excelio
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using JetBrains.Annotations;

    internal class ParameterValidator
    {
        /// <summary>
        /// Validates the specified parameters.
        /// </summary>
        /// <param name="parameters">The parameters.</param>
        /// <exception cref="ArgumentException">Missing input file.</exception>
        /// <exception cref="FileNotFoundException"></exception>
        public void Validate ([NotNull] IDictionary<string, object> parameters)
        {
            if (string.IsNullOrWhiteSpace((string) parameters[ArgumentParser.InputFile]))
                throw new ArgumentException("Missing input file.");

            var fileName = Path.GetFullPath((string) parameters[ArgumentParser.InputFile]);

            if (!File.Exists(fileName))
                throw new FileNotFoundException($"File '{(string)parameters[ArgumentParser.InputFile]}' not found.");
        }
    }
}