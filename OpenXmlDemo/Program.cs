// --------------------------------------------------------------------------------------
// <copyright file="program.cs" company="André Krämer - Software, Training & Consulting">
//      Copyright (c) 2014 André Krämer http://andrekraemer.de
// </copyright>
// <summary>
//  Open XML Demo Projekt
// </summary>
// --------------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlDemo
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string input;

            do
            {
                DisplayMenu();
                input = Console.ReadLine();
                ProcessInput(input);
            } while (input != null && !input.Equals("7", StringComparison.InvariantCultureIgnoreCase));
        }


        private static void DisplayMenu()
        {
            Console.Clear();
            Console.WriteLine("Open XML SDK Demo");
            Console.WriteLine("=================");
            Console.WriteLine("");
            Console.WriteLine("Bitte wählen Sie einen der folgenden Punkte aus:");
            Console.WriteLine("[1]: Autoreninfo aus Datei auslesen");
            Console.WriteLine("[2]: Neues Dokument erzeugen");
            Console.WriteLine("[3]: Plain Text aus Dokument auslesen");
            Console.WriteLine("[4]: Teilnehmerliste aus Vorlage 1 erzeugen");
            Console.WriteLine("[5]: Teilnehmerliste aus Vorlage 2 erzeugen");
            Console.WriteLine("[6]: Zertifikat erstellen");
            Console.WriteLine("[7]: Programm beenden");
        }

        private static void ProcessInput(string input)
        {
            const string fileName = @".\Hallo Basta.docx";

            switch (input)
            {
                case "1":
                    DisplayAuthor(fileName);
                    break;
                case "2":
                    CreateDocument();
                    break;
                case "3":
                    Console.WriteLine(ReadTextFromDocument(fileName));
                    break;
                case "4":
                    CreateAttendeeList();
                    break;
                case "5":
                    CreateAttendeeList2();
                    break;
                case "6":
                    MailMerge();
                    break;
                case "7":
                    break;
                default:
                    Console.WriteLine("Bitte eine gültige Auswahl wählen");
                    break;
            }
            Console.WriteLine("Bitte Eingabetaste drücken");
            Console.ReadLine();
        }


        /// <summary>
        /// Erstellt für jeden Teilnehmer ein Zertifikat auf Basis der Vorlage Zertifikat.docx
        /// </summary>
        private static void MailMerge()
        {
            const string template = @".\Zertifikat.docx";
            var training = GenerateSampleData();

            foreach (var attendee in training.Attendees)
            {
                var destinationFileName = template.Replace(".docx",
                    string.Format(" {0} {1} {2} {3}.docx", training.Title, attendee.FirstName, attendee.LastName,
                        training.From.ToString("yyyy-MM-dd")));
                File.Copy(template, destinationFileName, true);

                using (var document = WordprocessingDocument.Open(destinationFileName, true))
                {
                    string documentText;
                    using (var sr = new StreamReader(document.MainDocumentPart.GetStream()))
                    {
                        documentText = sr.ReadToEnd();
                    }

                    // "Felder" als Text suchen. Dies ist nicht sehr zuverlässig, da Word gerne einmal 
                    // unnötige Runs einfügt und somit das Feld nicht mehr vollständig gefunden werden kann
                    documentText = new Regex("SeminartitelFeld", RegexOptions.IgnoreCase).Replace(documentText,
                        training.Title);
                    documentText = new Regex("Punkt1Feld", RegexOptions.IgnoreCase).Replace(documentText,
                        training.Contents[0]);
                    documentText = new Regex("Punkt2Feld", RegexOptions.IgnoreCase).Replace(documentText,
                        training.Contents[1]);
                    documentText = new Regex("Punkt3Feld", RegexOptions.IgnoreCase).Replace(documentText,
                        training.Contents[2]);
                    documentText = new Regex("Punkt4Feld", RegexOptions.IgnoreCase).Replace(documentText,
                        training.Contents[3]);
                    documentText = new Regex("Punkt5Feld", RegexOptions.IgnoreCase).Replace(documentText,
                        training.Contents[4]);
                    documentText = new Regex("DatumFeld", RegexOptions.IgnoreCase).Replace(documentText,
                        DateTime.Today.ToShortDateString());

                    documentText = new Regex("AnredeFeld", RegexOptions.IgnoreCase).Replace(documentText, attendee.Title);
                    documentText = new Regex("VornameFeld", RegexOptions.IgnoreCase).Replace(documentText,
                        attendee.FirstName);
                    documentText = new Regex("NachnameFeld", RegexOptions.IgnoreCase).Replace(documentText,
                        attendee.LastName);

                    documentText = new Regex("VonFeld", RegexOptions.IgnoreCase).Replace(documentText,
    training.From.ToShortDateString());

                    documentText = new Regex("BisFeld", RegexOptions.IgnoreCase).Replace(documentText,
    training.To.ToShortDateString());


                    using (var sw = new StreamWriter(
                        document.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(documentText);
                    }
                    document.MainDocumentPart.Document.Save();
                    Console.WriteLine("Datei {0} im Programmordner erzeugt", destinationFileName);
                }
            }
        }

        /// <summary>
        /// Erzeugt eine Teilnehmerliste auf Basis der Vorlage Teilnerhmerliste1.docx
        /// Die Tabelle wird komplett im Code erzeugt
        /// </summary>
        private static void CreateAttendeeList()
        {
            const string template = @".\Teilnehmerliste1.docx";
            var training = GenerateSampleData();

            string destinationFileName = template.Replace("1",
                string.Format("{0} {1}", training.Title, training.From.ToString("yyyy-MM-dd")));
            File.Copy(template, destinationFileName, true);

            using (WordprocessingDocument document = WordprocessingDocument.Open(destinationFileName, true))
            {
                var docPart = document.MainDocumentPart;
                var doc = docPart.Document;

                var table = new Table();

                var borders = new TableBorders
                {
                    TopBorder = new TopBorder { Val = BorderValues.Single, Size = 6 },
                    LeftBorder = new LeftBorder { Val = BorderValues.Single, Size = 6 },
                    BottomBorder = new BottomBorder { Val = BorderValues.Single, Size = 6 },
                    RightBorder = new RightBorder { Val = BorderValues.Single, Size = 6 },
                    InsideHorizontalBorder = new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
                    InsideVerticalBorder = new InsideVerticalBorder { Val = BorderValues.Dashed, Size = 6 }
                };

                var tableProperties = new TableProperties();
                tableProperties.Append(borders);
                table.Append(tableProperties);

                var tableRow = new TableRow();

                var tableCell = CreateTableHeader("Titel");
                tableRow.Append(tableCell);

                tableCell = CreateTableHeader("Vorname");
                tableRow.Append(tableCell);

                tableCell = CreateTableHeader("Nachname");
                tableRow.Append(tableCell);


                tableCell = CreateTableHeader("Unterschrift");
                tableRow.Append(tableCell);
                table.Append(tableRow);

                foreach (Person attendee in training.Attendees)
                {
                    tableRow = new TableRow();

                    tableRow.Append(CreateTableCell(attendee.Title));
                    tableRow.Append(CreateTableCell(attendee.FirstName));
                    tableRow.Append(CreateTableCell(attendee.LastName));
                    tableRow.Append(CreateTableCell(""));
                    table.Append(tableRow);
                }

                doc.Body.Append(table);
                doc.Save();
            }

            Console.WriteLine("Datei {0} erzeugt", destinationFileName);
            Process.Start(destinationFileName);
        }

        /// <summary>
        ///  Erzeugt eine neue Überschriftentabellenzelle
        /// </summary>
        /// <param name="content">Spaltenüberschrift</param>
        /// <returns>Die Tabellenzelle</returns>
        private static TableCell CreateTableHeader(string content)
        {
            var text = new Text(content);
            var runProperties = new RunProperties();
            runProperties.Append(new Bold());
            var run = new Run();
            run.Append(runProperties);
            run.Append(text);
            var paragraph = new Paragraph(run);
            var tableCell = new TableCell(paragraph);

            var tcp = new TableCellProperties();
            var tcw = new TableCellWidth {Type = TableWidthUnitValues.Dxa, Width = "6000"};
            tcp.Append(tcw);
            tableCell.Append(tcp);
            return tableCell;
        }

        /// <summary>
        /// Erzeugt eine neue Tabellenzelle
        /// </summary>
        /// <param name="content">Der Inhalt der Zelle</param>
        /// <returns>Die Zelle</returns>
        private static TableCell CreateTableCell(string content)
        {
            var text = new Text(content);
            var run = new Run(text);
            var paragraph = new Paragraph(run);
            var tableCell = new TableCell(paragraph);

            var tcp = new TableCellProperties();
            var tcw = new TableCellWidth {Type = TableWidthUnitValues.Dxa, Width = "6000"};
            tcp.Append(tcw);
            tableCell.Append(tcp);
            return tableCell;
        }

        /// <summary>
        /// Erstellt eine Teilnehmerliste auf basis der Vorlage Teilnehmerliste2.docx
        /// Die Tabelle ist bereits in der Vorlage definiert und wird nur fortgeführt
        /// </summary>
        private static void CreateAttendeeList2()
        {
            const string template = @".\Teilnehmerliste2.docx";
            Training training = GenerateSampleData();

            string destinationFileName = template.Replace("2",
                string.Format(" {0} {1}", training.Title, training.From.ToString("yyyy-MM-dd")));
            File.Copy(template, destinationFileName, true);

            using (WordprocessingDocument document = WordprocessingDocument.Open(destinationFileName, true))
            {
                var docPart = document.MainDocumentPart;
                var doc = docPart.Document;


                // Erste Tabelle im Dokument suchen
                var table = doc.Body.Descendants<Table>().First();

                // Die letzte Zeile wird als Vorlage genutzt
                var templateRow = table.Elements<TableRow>().Last();

                foreach (var attendee in training.Attendees)
                {
                    // für jeden Teilnehmer die Vorlagenzeile clonen
                    var tableRow = templateRow.CloneNode(true) as TableRow;

                    // je Spalte erste den kompletten Inhalt löschen und dann den neuen Inhalt einfügen
                    tableRow.Descendants<TableCell>().ElementAt(0).RemoveAllChildren<Paragraph>();
                    tableRow.Descendants<TableCell>()
                        .ElementAt(0)
                        .Append(new Paragraph(new Run(new Text(attendee.Title))));

                    tableRow.Descendants<TableCell>().ElementAt(1).RemoveAllChildren<Paragraph>();
                    tableRow.Descendants<TableCell>()
                        .ElementAt(1)
                        .Append(new Paragraph(new Run(new Text(attendee.FirstName))));

                    tableRow.Descendants<TableCell>().ElementAt(2).RemoveAllChildren<Paragraph>();
                    tableRow.Descendants<TableCell>()
                        .ElementAt(2)
                        .Append(new Paragraph(new Run(new Text(attendee.LastName))));

                    table.Append(tableRow);
                }
                table.RemoveChild(templateRow);


                doc.Save();
            }

            Console.WriteLine("Datei {0} erzeugt", destinationFileName);
            Process.Start(destinationFileName);
        }

        /// <summary>
        /// List die Autoreninformationen aus einer Word Datei aus
        /// </summary>
        /// <param name="fileName">Dateiname einer docx Datei</param>
        private static void DisplayAuthor(string fileName)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, false))
            {
                string author = document.PackageProperties.Creator;
                Console.WriteLine(author);
            }
        }

        /// <summary>
        /// Erzeugt ein leeres Dokument mit dem Inhalt "Hallo Welt"
        /// </summary>
        private static void CreateDocument()
        {
            string fileName = Path.Combine(@".\", Path.GetRandomFileName() + ".docx");

            using (
                WordprocessingDocument document = WordprocessingDocument.Create(fileName,
                    WordprocessingDocumentType.Document))
            {
                var text = new Text("Hallo Welt");
                var run = new Run(text);
                var paragraph = new Paragraph(run);
                var body = new Body(paragraph);
                var doc = new Document(body);

                document.AddMainDocumentPart();

                document.MainDocumentPart.Document = doc;

                document.MainDocumentPart.Document.Save();

            }

            Console.WriteLine("Datei {0} erzeugt", fileName);
            Process.Start(fileName);
        }


        /// <summary>
        /// List den kompletten Text einer Datei als Plain Text aus
        /// </summary>
        /// <param name="fileName">Name der Datie</param>
        /// <returns>Den Text</returns>
        public static string ReadTextFromDocument(string fileName)
        {
            string results = null;

            using (var document = WordprocessingDocument.Open(fileName, false))
            {
                var docPart = document.MainDocumentPart;
                results = GetPlainText(docPart.Document.Body);

            }

            return results;
        }

        /// <summary>
        /// Extrahiert den Text aus einem XML Element
        /// </summary>
        /// <param name="rootElement">Das Element</param>
        /// <param name="sb">Optional: ein Stringbuilder</param>
        /// <returns>Den Text</returns>
        public static string GetPlainText(OpenXmlElement rootElement, StringBuilder sb = null)
        {
            if (sb == null)
            {
                sb = new StringBuilder();
            }

            foreach (var childElement in rootElement.Elements())
            {
                switch (childElement.LocalName)
                {
                    case "t": // Text
                        sb.Append(childElement.InnerText);
                        break;
                    case "tab": // Tab 
                        sb.Append("\t");
                        break;
                    case "cr": // Zeilenumbruch
                    case "br": // Seitenumbruch
                        sb.Append(Environment.NewLine);
                        break;
                    case "p":// Absatz 
                        GetPlainText(childElement,sb);
                        sb.AppendLine(Environment.NewLine);
                        break;

                    default:
                        GetPlainText(childElement, sb);
                        break;
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Legt Demo Daten an
        /// </summary>
        /// <returns>Demo Daten</returns>
        public static Training GenerateSampleData()
        {
            var training = new Training
            {
                Title = "OpenXML SDK",
                From = DateTime.Today.AddDays(-2),
                To = DateTime.Today.AddDays(-1)
            };

            training.Contents.Add("Überblick Open XML SDK");
            training.Contents.Add("Lesen von Dokumenteigenschaften");
            training.Contents.Add("Erstellen von neuen Dokumenten");
            training.Contents.Add("Lesen von bestehenden Dokumenten");
            training.Contents.Add("Verändern bestehender Dokumente");

            training.Attendees.Add(new Person { FirstName = "Wilhelm", LastName = "Brause", Title = "Herr" });
            training.Attendees.Add(new Person { FirstName = "Peter", LastName = "Schmitz", Title = "Herr" });
            training.Attendees.Add(new Person { FirstName = "Laura", LastName = "Buitoni", Title = "Frau" });


            return training;
        }
    }
}