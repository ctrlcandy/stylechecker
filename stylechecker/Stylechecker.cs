using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using Xceed.Words.NET;
using Font = DocumentFormat.OpenXml.Wordprocessing.Font;

namespace stylechecker
{
    public class Stylechecker
    {
        private int DefaultOpenXmlFontSize;
        private string DefaultFont { get; set; } = string.Empty;
        private int DefaultFontSize;
        private double DefaultLineSpacing;
        private string DefaultAlignment { get; set; } = string.Empty;

        private Fonts fontTable;
        public string ResultErrors { get; set; } = string.Empty;
        public string ResultWarnings { get; set; } = string.Empty;

        public Stylechecker(string Font, int FontSize, double LineSpacing, string Alignment)
        {
            DefaultFont = Font;
            DefaultFontSize = FontSize;
           
            switch (Alignment)
            {
                case ("По центру"):
                    DefaultAlignment = "center";
                    break;
                case ("По левому краю"):
                    DefaultAlignment = "left";
                    break;
                case ("По правому краю"):
                    DefaultAlignment = "right";
                    break;
                default:
                    DefaultAlignment = "both";
                    break;
            }
            DefaultLineSpacing = LineSpacing * 12;
        }

        public void GetFontInfo(uint line, string text, Xceed.Document.NET.Paragraph p)
        {
            HashSet<string> usingFonts = new HashSet<string>();
            try
            {
                foreach (var magic in p.MagicText)
                {
                    if (magic.formatting.FontFamily.Name != DefaultFont)
                    {
                        if (usingFonts.Contains(magic.formatting.FontFamily.Name) == false)
                        {
                            usingFonts.Add(magic.formatting.FontFamily.Name);
                            ResultErrors += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                        " - используется неверный шрифт: " + magic.formatting.FontFamily.Name + '\n';
                        }
                    }
                    else if (magic.formatting.FontFamily == null & usingFonts.Contains("null") == false)
                    {
                        usingFonts.Add("null");

                        ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                   " - предположительно используется неверный шрифт. Возможные варианты:" + '\n';
                        foreach (var f in fontTable.ChildElements)
                        {
                            Font font = f as Font;
                            if (!font.Name.ToString().Contains(DefaultFont) & !font.Name.ToString().Contains("serif"))
                                ResultWarnings += "* " + font.Name + '\n';
                        }
                    }
                }
                usingFonts.Clear();
            }
            catch
            {
                ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                    " - предположительно используется неверный шрифт. Возможные варианты:" + '\n';
                foreach (var f in fontTable.ChildElements)
                {
                    Font font = f as Font;
                    if (!font.Name.ToString().Contains(DefaultFont) & !font.Name.ToString().Contains("serif"))
                        ResultWarnings += "* " + font.Name + '\n';
                }
            }
        }

        public void GetFontSizeInfo(uint line, string text, Xceed.Document.NET.Paragraph p)
        {
            HashSet<double?> usingSize = new HashSet<double?>();
            try
            {
                foreach (var magic in p.MagicText)
                {
                    if ((magic.formatting.Size == null || magic.formatting.Size == 0) & (usingSize.Contains(null) == false || usingSize.Contains(0) == false))
                    {
                        usingSize.Add(0);
                        if (DefaultFontSize != 0 & DefaultOpenXmlFontSize != 0)
                        {
                            ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            $" - предположительно используется неверный кегль. Возможно: {DefaultOpenXmlFontSize/2}" + '\n';
                        }
                        else
                        {
                            ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            " - предположительно используется неверный кегль. " + '\n';
                        }
                    }
                    else if (magic.formatting.Size != DefaultFontSize & usingSize.Contains(magic.formatting.Size) == false)
                    {
                        usingSize.Add(magic.formatting.Size);
                        ResultErrors += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            " - используется неверный кегль: " + magic.formatting.Size + '\n';
                    }
                }
                usingSize.Clear();
            }
            catch
            {
                ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                        " - предположительно используется неверный кегль. " + '\n';
            }
        }

        public void GetLineSpacingInfo(uint line, string text, Xceed.Document.NET.Paragraph p)
        {
            try
            {
                if (p.LineSpacing == 0)
                {
                    throw new Exception();
                }
                else if (p.LineSpacing!= DefaultLineSpacing)
                {
                    ResultErrors += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                        " - используется неверный междустрочный интервал: " +
                        (p.LineSpacing / 12).ToString("0.00").Replace(',', '.') + $" вместо {(DefaultLineSpacing / 12).ToString("0.00").Replace(',', '.')} строки" + '\n';
                }
            }
            catch
            {
                ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                    " - предположительно используется неверный междустрочный интервал." + '\n';
            }
        }

        public void GetAlignmentInfo(uint line, string text, Xceed.Document.NET.Paragraph p)
        {
            try
            {
                if (p.Alignment.ToString() != DefaultAlignment)
                {
                    if (p.Alignment.ToString() == "center")
                    {
                        ResultErrors += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            " - используется выравнивание по центру." + '\n';
                    }
                    else if (p.Alignment.ToString() == "left")
                    {
                        ResultErrors += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            " - используется выравнивание по левому краю." + '\n';
                    }
                    else if (p.Alignment.ToString() == "right")
                    {
                        ResultErrors += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            " - используется выравнивание по правому краю." + '\n';
                    }
                    else if (p.Alignment.ToString() == "both")
                    {
                        ResultErrors += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            " - используется выравнивание по ширине." + '\n';
                    }
                    else
                    {
                        ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            " - предположительно используется неверное выравнивание." + '\n';
                    }
                }
            }
            catch { }
        }


        void SetDefaults(string filePath)
        {
        
                using (WordprocessingDocument docx = WordprocessingDocument.Open(filePath, true))
                {
                    DocDefaults defaults = docx.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults;
                    RunFonts runFont = defaults.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts;
                    try
                    {
                        DefaultOpenXmlFontSize = Convert.ToInt32(defaults.RunPropertiesDefault.RunPropertiesBaseStyle.FontSize.Val.Value);
                    }
                    catch
                    {
                        DefaultOpenXmlFontSize = 0;
                    }

                    fontTable = docx.MainDocumentPart.FontTablePart.Fonts;
                    docx.Close();
                }
            
        }

        public void MyDocument(string filePath, bool checkFonts = true, bool checkFontSize = true,
            bool checkLineSpacing = true, bool checkAlignment = true)
        {
            try {
                SetDefaults(filePath);
                if (File.Exists(filePath))
                {
                    using (var doc = DocX.Load(filePath))
                    {
                        uint line = 0;
                        string text;

                        foreach (var p in doc.Paragraphs)
                        {
                            line++;  

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

                            if (checkFonts)
                            {
                                GetFontInfo(line, text, p);
                            }
                            if (checkFontSize)
                            {
                                GetFontSizeInfo(line, text, p);
                            }
                            if (checkLineSpacing)
                            {
                                GetLineSpacingInfo(line, text, p);
                            }
                            if (checkAlignment)
                            {
                                GetAlignmentInfo(line, text, p);
                            }
                        }
                    }
                }
            }
            catch
            {
                ResultErrors = ResultWarnings = "Выбран некорректный тип файла!" + '\n' +
                    "Убедитесь, что выбран файл формата DOCX. " +
                    "При необоходимости откройте Microsoft Word и пересохраните Ваш документ в нужном формате.";
            }
        }
    }
}