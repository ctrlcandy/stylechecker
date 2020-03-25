using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xceed.Words.NET;


namespace ConsoleApp2
{
    class Program
    {
        private static string DefaultFontName { get; set; }
        private static int DefaultFontSize { get; set; }
        public string Name { get; }

        static void Main(string[] args)
        {
            using (WordprocessingDocument docx = WordprocessingDocument.Open("KMBO.docx", true))
            {
                Fonts fontTable = docx.MainDocumentPart.FontTablePart.Fonts;
                docx.Close();

                using (var doc = DocX.Load("KMBO.docx"))
                {
                    uint line = 0;
                    string text;
                    HashSet<string> usingFonts = new HashSet<string>();

                    foreach (var p in doc.Paragraphs)
                    {
                        line++; // строка документа  

                        if (p.Text.Length > 15)
                        {
                            text = p.Text.Remove(15).TrimStart(' ').TrimStart('\t').TrimStart('\n') + "...";
                            if (text == "...")
                                continue;
                        }
                        else
                        {
                            text = p.Text.TrimStart(' ').TrimStart('\t').TrimStart('\n');
                            if (text == "")
                                continue;
                        }

                        // получение шрифтов, работает сносно
                        try
                        {
                            foreach (var magic in p.MagicText)
                            { 
                                if (magic.formatting.FontFamily.Name != "Times New Roman")
                                {
                                    if (usingFonts.Contains(magic.formatting.FontFamily.Name) == false)
                                    {
                                        usingFonts.Add(magic.formatting.FontFamily.Name);
                                        Console.WriteLine("Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                    " - используется неверный шрифт: " + magic.formatting.FontFamily.Name);
                                    }
                                } else if (magic.formatting.FontFamily == null & usingFonts.Contains("null") == false)
                                {
                                    usingFonts.Add("null");

                                    Console.WriteLine("Строка " + line + " (начинается со слов \"" + text + "\" )" +
                               " - предположительно используется неверный шрифт. Возможные варианты:");
                                    foreach (var f in fontTable.ChildElements)
                                    {
                                        Font font = f as Font;
                                        if (!font.Name.ToString().Contains("Times New Roman") & !font.Name.ToString().Contains("serif"))
                                            Console.WriteLine("* " + font.Name);
                                    }
                                }
                            }
                            usingFonts.Clear();
                        }
                        catch
                        {
                            Console.WriteLine("Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                " - предположительно используется неверный шрифт. Возможные варианты:");
                            foreach (var f in fontTable.ChildElements)
                            {
                                Font font = f as Font;
                                if (!font.Name.ToString().Contains("Times New Roman" ) & !font.Name.ToString().Contains("serif"))
                                    Console.WriteLine("* " + font.Name);
                            }
                        }

                        //получение кегля, работает плюс-минус нормально
                        try
                        {
                            foreach (var magic in p.MagicText)
                            {
                                if (magic.formatting.Size == null)
                                {
                                    continue;
                                }
                                else if (magic.formatting.Size != 14)
                                {
                                    Console.WriteLine("Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                        " - используется неверный кегль: " + magic.formatting.Size);
                                }
                            }
                        }
                        catch { }

                        //получение междустрочного интервала, в целом работает корректно
                        try
                        {
                            if (p.LineSpacing != 18)
                            {
                                Console.WriteLine("Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                    " - используется неверный междустрочный интервал: " + (p.LineSpacing / 12).ToString("0.00").Replace(',', '.') + " вместо 1.5 строки");
                            }
                        }
                        catch (NullReferenceException) {
                            Console.WriteLine("Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                " - предположительно используется неверный междустрочный интервал.");
                        }
                        catch (InvalidOperationException) {
                            Console.WriteLine("Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                " - предположительно используется неверный междустрочный интервал.");
                        }

                        //получение информации о выравнивании, работает вроде как корректно 
                        try
                        {
                            if (p.Alignment.ToString() != "both")
                            {
                                if (p.Alignment.ToString() == "center")
                                {
                                    Console.WriteLine("Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                        " - используется выравнивание по центру вместо выравнивания по ширине.");
                                }
                                else if (p.Alignment.ToString() == "left")
                                {
                                    Console.WriteLine("Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                        " - используется выравнивание по левому краю вместо выравнивания по ширине.");
                                }
                                else if (p.Alignment.ToString() == "right")
                                {
                                    Console.WriteLine("Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                        " - используется выравнивание по правому краю вместо выравнивания по ширине.");
                                }
                                else
                                {
                                    Console.WriteLine("Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                        " - предположительно используется неверное выравнивание по ширине");
                                }
                            }
                        }
                        catch { }
                    }
                }
                Console.ReadKey();
            }
        }
    }
}