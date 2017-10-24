namespace excelio
{
    using GemBox.Spreadsheet;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Security.Cryptography;
    using System.Text;
    using Extensions;
    using JetBrains.Annotations;
    using LiteDB;
    using Readers;

    internal static class LiteDbExtensions
    {
        public static LiteCollection<T> GetCollection<T> ([NotNull] this LiteDatabase db)
        {
            return db.GetCollection<T>($"{typeof(T).Name}s");
        }

        public static void Shuffle<T> ([NotNull] this IList<T> list)
        {
            var provider = new RNGCryptoServiceProvider();
            var n = list.Count;
            while (n > 1)
            {
                var box = new byte[1];
                do
                {
                    provider.GetBytes(box);
                }
                while (!(box[0] < n * (byte.MaxValue / n)));

                var k = (box[0] % n);
                n--;

                var value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }
    }

    internal class Program
    {
        static void Initialize ()
        {
            // Set license key to use GemBox.Spreadsheet in a Free mode.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            // If sample exceeds Free version limitations then continue as trial version:
            // https://www.gemboxsoftware.com/Spreadsheet/help/html/Evaluation_and_Licensing.htm
            SpreadsheetInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
        }

        private static void Main (string[] args)
        {
            /*
             * excelio --file=<file>.xls(x|m) 
             */
            var parser = new ArgumentParser();
            var parameters = parser.Parse(args);

            try
            {
                var validator = new ParameterValidator();
                validator.Validate(parameters);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            Initialize();

            var folder = Path.GetDirectoryName(parameters.GetFullExcelFilePath());
            var files = Directory.GetFiles(folder, "*.xlsx");

            var outputFolder = parameters.GetOutputPath();
            if (Directory.Exists(outputFolder))
            {
                Directory.Delete(outputFolder, true);
            }

            foreach (var file in files)
            {
                var workbook = ExcelFile.Load(file);
                var reader = ExcelReaderFactory.Create((int) parameters[ArgumentParser.FileFormat], workbook);

                if (null != reader)
                    PersistParticipantInfo(reader, outputFolder);
            }

            CreateExportFiles(outputFolder, parameters.ShouldRandomize());

            Console.WriteLine("All files read. Press any key to exit!");
            Console.ReadLine();
        }

        private static void PersistParticipantInfo ([NotNull] ExcelReader reader, [NotNull] string outputFolder)
        {
            reader.Read();

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            var dbFile = Path.Combine(outputFolder, "participation.db");

            using (var db = new LiteDatabase(dbFile))
            {
                var col = db.GetCollection<Participant>();
                col.Insert(reader.Participants);
            }
        }

        private static void CreateExportFiles ([NotNull] string outputFolder, bool shouldRandomize)
        {
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            var dbFile = Path.Combine(outputFolder, "participation.db");

            using (var db = new LiteDatabase(dbFile))
            {
                var participations = db.GetCollection<Participant>();
                var events = GetEvents(participations);

                foreach (var @event in events)
                {
                    var file = Path.Combine(outputFolder, $"{@event}.csv");

                    var participants = participations
                        .Find(p => p.Event == @event)
                        .ToList();

                    if (shouldRandomize)
                    {
                        participants.Shuffle();
                        //var randomizer = new Randomizer<Participant>();

                        //// randomize the list twice
                        //var randomized = randomizer.Randomize(
                        //    randomizer.Randomize(participants)
                        //    );

                        //participants = randomized;
                    }

                    var ef = SaveParticipantsToExcel(participants);
                    CsvStorer.SaveCvs(ef, file);
                }
            }
        }

        private static ExcelFile SaveParticipantsToExcel ([NotNull] List<Participant> participants)
        {
            // write to excel file and save as csv
            var ef = new ExcelFile();
            var ws = ef.Worksheets.Add("Participants");

            // add headers
            var headers = new List<string>
            {
                "Member ID",
                "Last Name",
                "First Name",
                "Date of Birth",
                "Gender",
                "Club Number",
                "Club Name",
                "Country",
                "Region",
                "Number",
                "Category",
                "Team"
            };

            for (var col = 0; col < headers.Count; col++)
            {
                ws.Cells[0, col].Value = headers[col];
            }

            // add members
            var row = 1;
            var number = 1;
            participants
                .ForEach(m =>
                {
                    ws.Cells[row, 0].Value = m.Id;
                    ws.Cells[row, 1].Value = m.Name;
                    ws.Cells[row, 2].Value = "";
                    ws.Cells[row, 3].Value = m.YearOfBirth ?? 1900;
                    ws.Cells[row, 4].Value = m.Event.Contains("pojat") ? "M" : "F";
                    ws.Cells[row, 5].Value = Clubs.GetIdentifier(m.Club);
                    ws.Cells[row, 6].Value = m.Club;
                    ws.Cells[row, 7].Value = "FIN";
                    ws.Cells[row, 8].Value = "";
                    ws.Cells[row, 9].Value = number;
                    ws.Cells[row, 10].Value = GetMappedCategoryName(m.Event);
                    ws.Cells[row, 11].Value = m.Team;

                    row++;
                    number++;
                });

            return ef;
        }

        private static List<string> GetEvents ([NotNull] LiteCollection<Participant> participants)
        {
            return participants
                .FindAll()
                .Select(p => p.Event)
                .ToList();
        }

        private static string GetMappedCategoryName ([NotNull] string eventName)
        {
            return categoryMapper[eventName];
        }

        private static Dictionary<string, string> categoryMapper = new Dictionary<string, string>
        {
            // name of event        -- Category
            { "1 lk TRA yksilö pojat", "Luokka 1 Pojat" },
            { "1 lk TRA yksilö tytöt", "Luokka 1 Tytöt" },
            { "2 lk TRA yksilö pojat", "Luokka 2 Pojat" },
            { "2 lk TRA yksilö tytöt", "Luokka 2 Tytöt" },
            { "3 lk TRA yksilö pojat", "Luokka 3 Pojat" },
            { "3 lk TRA yksilö tytöt", "Luokka 3 Tytöt" },

            { "1 lk SYNKRO pojat", "Luokka 1 Pojat" },
            { "1 lk SYNKRO tytöt", "Luokka 1 Tytöt" },
            { "2 lk SYNKRO pojat", "Luokka 2 Pojat" },
            { "2 lk SYNKRO tytöt", "Luokka 2 Tytöt" },

            // WAGC
            { "13-14 yksilö pojat", "WAGC 13-14 TRI Pojat" },
            { "13-14 yksilö tytöt", "WAGC 14-14 TRI Tytöt" },
            { "15-16 yksilö pojat", "WAGC 15-16 TRI Pojat" },
            { "15-16 yksilö tytöt", "WAGC 15-16 TRI Tytöt" },
            { "13-14 synkro tytöt", "WAGC 13-14 TRS Tytöt" },
            { "17-21 yksilö pojat", "Luokka 4 Seniorit Miehet" },
            { "17-21 yksilö tytöt", "Luokka 4 Seniorit Naiset" },
            { "17-21 synkro tytöt", "Luokka 4 Seniorit Naiset" },

            // DMT
            { "13-14 DMT pojat", "Nuoret alle 14v Pojat" },
            { "13-14 DMT tytöt", "Nuoret 14v Tytöt" },
            { "15-16 DMT pojat", "Juniorit 15-16v Pojat" },
            { "15-16 DMT tytöt", "Juniorit 15-16v Tytöt" },
            { "17- DMT pojat", "Seniorit 17+ Pojat" },
            { "17- DMT tytöt", "Seniorit 17+ Tytöt" },
            { "alle 12 DMT pojat", "Pojat alle 12v" },
            { "alle 12 DMT tytöt", "Tytöt alle 12v" },

            { "Tuomarit tytöt", "Tuomarit" }
        };
    }

    internal class CsvStorer
    {
        public static void SaveCvs ([NotNull] ExcelFile ef, [NotNull] string file)
        {
            var options = new CsvSaveOptions(CsvType.SemicolonDelimited);
            ef.Save(file, options);

            ConvertToLatin1(file);
        }

        /// <summary>
        /// Converts the specified file name in place to have iso-8859-1 (latin 1) character set.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        private static void ConvertToLatin1 ([NotNull] string fileName)
        {
            var input = File.ReadAllBytes(fileName);
            var output = Encoding.Convert(Encoding.UTF8, Encoding.GetEncoding("iso-8859-1"), input);
            File.WriteAllBytes(fileName, output);
        }
    }

    internal class Randomizer<T>
    {
        [Pure, NotNull]
        public List<T> Randomize ([NotNull] List<T> source)
        {
            var clone = new List<T>(source);
            var randomized = new List<T>();

            var random = new Random();
            while (clone.Count > 0)
            {
                var index = random.Next(clone.Count);
                var entry = clone[index];
                randomized.Add(entry);

                clone.Remove(entry);
            }

            return randomized;
        }
    }
}
