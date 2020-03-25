using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xceed.Words.NET;

namespace ConsoleApp
{
    class Program
    {
        private static string DefaultFontName { get; set; }
        private static int DefaultFontSize { get; set; }

        static void Main(string[] args)
        {

            using (var doc = DocX.Load("tasks.docx"))
            {
                uint line = 0;
                
                foreach (var p in doc.Paragraphs)
                {
                    line++; // строка документа
                  	Console.WriteLine("***");

                    // получение шрифтов, работает очень не очень корректно )))
                    try
                    {
                        if (p.MagicText.FirstOrDefault().formatting.FontFamily.Name != "Times New Roman")
                        {
                            Console.WriteLine("Строка " + line + " - используется неверный шрифт: " + p.MagicText.FirstOrDefault().formatting.FontFamily.Name);
                        }
                    }
                    catch (NullReferenceException) {
                        Console.WriteLine("Строка " + line + " - предположительно используется неверный шрифт.");
                    }

                    //получение кегля, работает не очень корректно
                    try
                    {
                        if (p.MagicText.First().formatting.Size != 14)
                        {
                            Console.WriteLine("Строка " + line + " - используется неверный кегль: " + p.MagicText.First().formatting.Size);
                        }
                    }
                    catch (NullReferenceException) {
                        Console.WriteLine("Строка " + line + " - предположительно используется неверный кегль.");
                    }
                    catch (InvalidOperationException) {
                        Console.WriteLine("Строка " + line + " - предположительно используется неверный кегль.");
                    }

                    //получение междустрочного интервала, в целом работает корректно, но были найдены случаи, когда не особо
                    try
                    {
                        if (p.LineSpacing != 18)
                        {
                            Console.WriteLine("Строка " + line + " - используется неверный междустрочный интервал: " + (p.LineSpacing/12).ToString("0.00").Replace(',','.') + " вместо 1.5 строки");
                        }
                    }
                    catch (NullReferenceException) {
                        Console.WriteLine("Строка " + line + " - предположительно используется неверный междустрочный интервал.");
                    }
                    catch (InvalidOperationException) {
                        Console.WriteLine("Строка " + line + " - предположительно используется неверный междустрочный интервал.");
                    }

                    //получение информации о выравнивании, предположительно работает корректно 
                    try
                    {
                        if (p.Alignment.ToString() != "both")
                        {
                            if (p.Alignment.ToString() == "center")
                            {
                                Console.WriteLine("Строка " + line + " - используется выравнивание по центру вместо выравниыания по ширине.");
                            }
                            else if (p.Alignment.ToString() == "left")
                            {
                                Console.WriteLine("Строка " + line + " - используется выравнивание по левому краю вместо выравниыания по ширине.");
                            }
                            else if (p.Alignment.ToString() == "right")
                            {
                                Console.WriteLine("Строка " + line + " - используется выравнивание по правому краю вместо выравниыания по ширине.");
                            }
                            else
                            {
                                Console.WriteLine("Строка " + line + " - предположительно используется неверное выравнивание по ширине");
                            }
                        }
                    }
                    catch (NullReferenceException) { }


                }
            }
            Console.WriteLine("Done");
            Console.ReadKey();
        }
    }