namespace excelio.Readers
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Diagnostics.CodeAnalysis;
    using System.Linq;
    using System.Text.RegularExpressions;
    using GemBox.Spreadsheet;
    using JetBrains.Annotations;
    using LiteDB;

    internal static class ReaderExtensions
    {
        private static string FirstCellValue ([NotNull] this ExcelRow row)
        {
            var cells = row.AllocatedCells;

            if (null == cells || 0 != cells[0].Column.Index)
                return null;

            return cells[0].StringValue;
        }

        public static bool IsEventRow ([NotNull] this ExcelRow row, out string eventName)
        {
            eventName = null;

            var value = row.FirstCellValue();

            var isEvent = !string.IsNullOrWhiteSpace(value)
                          && (value.Contains("pojat") || value.Contains("tytöt"));

            if (isEvent)
            {
                eventName = value;
            }

            return isEvent;
        }

        public static bool IsDisciplineRow ([NotNull] this ExcelRow row, out string discipline)
        {
            discipline = null;

            var value = row.FirstCellValue();
            var isDisciplineRow = !string.IsNullOrWhiteSpace(value)
                                  && (value.Contains("Tupla Minitrampoliini") || value.Contains("Trampoliini"));

            discipline = value;

            return isDisciplineRow;
        }

        public static bool IsSkippableRow ([NotNull] this ExcelRow row)
        {
            var value = row.FirstCellValue();

            var isSkippable = string.IsNullOrWhiteSpace(value)                  // empty
                              || value.Contains("/") || value.Contains(" ja ");   // syncro

            return isSkippable;
        }

        private static string GetSurname ([NotNull] this string fullName)
        {
            var elements = fullName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return elements[0].Trim();
        }

        public static List<Participant> GetParticipants ([NotNull] this ExcelRow row, [NotNull] string club)
        {
            var participants = new List<Participant>();

            var cells = row.AllocatedCells;

            if (0 == cells.Count)
                return participants;

            foreach (var cell in cells)
            {
                const string pattern = @"[^\(\);]+";
                var matches = Regex.Matches(cell.StringValue, pattern);

                // only 1 match - single
                // 4 matches (name, y of b, name, y of b) - syncro
                if (1 == matches.Count)
                {
                    var p = new Participant
                    {
                        Name = cell.StringValue.Trim(),
                        Club = club,
                        Team = string.Empty
                    };

                    participants.Add(p);
                }
                else if (4 == matches.Count)
                {
                    var team = $"{matches[0].Value.GetSurname()} ({matches[1].Value}) / {matches[2].Value.GetSurname()} ({matches[3].Value})";

                    var p1 = new Participant
                    {
                        Name = team,
                        Club = club,
                        Team = team
                    };

                    participants.Add(p1);
                }
            }

            return participants;
        }
    }

    [SuppressMessage ("ReSharper", "InconsistentNaming")]
    internal class Clubs
    {
        private static readonly IReadOnlyCollection<string> clubs = new ReadOnlyCollection<string>(new List<string>
        {
            string.Empty,
            "Bounce",
            "EsTT",
            "Fliku-82",
            "LTV",
            "TV",
            "VSH"
        });

        public static int GetIdentifier ([NotNull] string name)
        {
            var index = 0;

            foreach (var club in clubs)
            {
                if (string.Equals(club, name, StringComparison.InvariantCultureIgnoreCase))
                {
                    return index;
                }

                index++;
            }

            return -1;
        }
    }


    internal class Membership
    {
        [BsonId(true)]
        public int Id { get; set; }

        public string Name { get; set; }

        public int ClubId { get; set; }

        public string Club { get; set; }

        public int YearOfBirth { get; set; }

        public string Gender { get; set; }
    }

    internal class Participant
    {
        [BsonId (true)]
        public int Id { get; set; }

        public string Name { get; set; }

        public int? YearOfBirth { get; set; }

        public string Club { get; set; }

        public string Team { get; set; }

        public string Discipiline { get; set; }

        public string Event { get; set; }
    }

    internal class InputExcelReader : ExcelReader
    {
        public InputExcelReader ([NotNull] ExcelFile workbook) 
            : base(workbook)
        {
        }

        public override void Read ()
        {
            /*
             * Start reading from cell A1
             * Line represents event, if contains 'pojat' or 'tytöt'
             * Line can be skipped if contains '/' friendly syncro pair name
             * Syncro team members are separated with ';'
             * Syncro team member birth year is within parenthesis '(yyyy)'.
             * 
             * */
            
            // Read participation information 1st worksheet
            var ws = Workbook.Worksheets[0];
            var club = ws.Name;

            string discipline = null;
            string @event = null;

            foreach (var row in ws.Rows)
            {
                string result;

                if (row.IsDisciplineRow(out result))
                {
                    discipline = result;
                }
                else if (row.IsEventRow(out result))
                {
                    @event = result;
                }
                else if (row.IsSkippableRow())
                {
                }
                else
                {
                    // this row contains participant information
                    var p = row.GetParticipants(club);

                    if (!p.Any())
                        continue;

                    p.ForEach(e =>
                    {
                        e.Discipiline = discipline;
                        e.Event = @event;
                    });

                    Participants.AddRange(p);
                }
            }
        }
    }
}