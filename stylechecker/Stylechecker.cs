using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
        //private DocumentFormat.OpenXml.OpenXmlElementList styles;
        public string ResultErrors { get; set; } = string.Empty;
        public string ResultWarnings { get; set; } = string.Empty;

        public Process AssignedProcess { get; set; }
        private bool checkCopy;
        private bool checkErrors;
        private bool checkWarnings;

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
                    /*foreach (var s in styles)
                    {
                       // some shit
                    }*/

                    if (magic.formatting.FontFamily.Name != DefaultFont)
                    {
                        if (usingFonts.Contains(magic.formatting.FontFamily.Name) == false)
                        {
                            usingFonts.Add(magic.formatting.FontFamily.Name);
                            ResultErrors += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                        " - используется неверный шрифт: " + magic.formatting.FontFamily.Name + '\n';
                            if (checkCopy)
                            {
                                if (checkErrors)
                                {
                                    p.InsertText($"\n--- Используется неверный шрифт: {magic.formatting.FontFamily.Name}. ---");
                                }
                            }
                        }
                    }
                    else if (magic.formatting.FontFamily == null & usingFonts.Contains("null") == false)
                    {
                        usingFonts.Add("null");

                        ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                   " - предположительно используется неверный шрифт. Возможные варианты:" + '\n';

                        if (checkCopy)
                        {
                            if (checkWarnings)
                            {
                                p.InsertText("\n--- Предположительно используется неверный шрифт. Возможные варианты: ---");
                            }
                        }

                        foreach (var f in fontTable.ChildElements)
                        {
                            Font font = f as Font;
                            if (!font.Name.ToString().Contains(DefaultFont) & !font.Name.ToString().Contains("serif"))
                                ResultWarnings += "* " + font.Name + '\n';

                            if (checkCopy)
                            {
                                if (checkWarnings)
                                {
                                    p.InsertText($"\n* {font.Name}");
                                }
                            }
                        }
                    }
                }
                usingFonts.Clear();
            }
            catch
            {
                ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                    " - предположительно используется неверный шрифт. Возможные варианты:" + '\n';

                if (checkCopy)
                {
                    if (checkWarnings)
                    {
                        p.InsertText("\n--- Предположительно используется неверный шрифт. Возможные варианты: ---");
                    }
                }

                foreach (var f in fontTable.ChildElements)
                {
                    Font font = f as Font;
                    if (!font.Name.ToString().Contains(DefaultFont) & !font.Name.ToString().Contains("serif"))
                        ResultWarnings += "* " + font.Name + '\n';

                    if (checkCopy)
                    {
                        if (checkWarnings)
                        {
                            p.InsertText($"\n* {font.Name}");
                        }

                    }
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
                    if (magic.formatting.Size != DefaultFontSize & usingSize.Contains(magic.formatting.Size) == false & magic.formatting.Size != null & magic.formatting.Size != 0)
                    {
                        usingSize.Add(magic.formatting.Size);
                        ResultErrors += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            " - используется неверный кегль: " + magic.formatting.Size + '\n';
                        if (checkCopy)
                        {
                            if (checkErrors)
                            {
                                p.InsertText($"\n--- Используется неверный кегль: {magic.formatting.Size} ---");
                            }
                        }
                        usingSize.Add(magic.formatting.Size);
                    }
                    else if ((magic.formatting.Size == null || magic.formatting.Size == 0) & usingSize.Contains(DefaultOpenXmlFontSize / 2) == false)
                    {
                        if (DefaultFontSize != 0 & DefaultOpenXmlFontSize != 0 & (DefaultOpenXmlFontSize / 2) != DefaultFontSize)
                        {
                            ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            $" - предположительно используется неверный кегль. Возможно: {DefaultOpenXmlFontSize / 2}" + '\n';
                            if (checkCopy)
                            {
                                if (checkWarnings)
                                {
                                    p.InsertText($"\n--- Предположительно используется неверный кегль. Возможно: {DefaultOpenXmlFontSize / 2} ---");
                                }
                            }
                            usingSize.Add(DefaultOpenXmlFontSize / 2);
                        }
                        else if (usingSize.Contains(0) == false)
                        {
                            ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            " - предположительно используется неверный кегль. " + '\n';
                            if (checkCopy)
                            {
                                if (checkWarnings)
                                {
                                    p.InsertText("\n--- Предположительно используется неверный кегль. ---");
                                }
                            }
                            usingSize.Add(0);
                        }
                    }
                }
                usingSize.Clear();
            }
            catch
            {
                ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                        " - предположительно используется неверный кегль. " + '\n';
                if (checkCopy)
                {
                    if (checkWarnings)
                    {
                        p.InsertText("\n--- Предположительно используется неверный кегль. ---");
                    }
                }
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
                else if (Math.Round(p.LineSpacing, 2, MidpointRounding.AwayFromZero) != Math.Round(DefaultLineSpacing, 2, MidpointRounding.AwayFromZero))
                {
                    ResultErrors += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                        " - используется неверный междустрочный интервал: " +
                        (p.LineSpacing / 12).ToString("0.00").Replace(',', '.') + $" вместо {(DefaultLineSpacing / 12).ToString("0.00").Replace(',', '.')} строки" + '\n';
                    if (checkCopy)
                    {
                        if (checkErrors)
                        {
                            p.InsertText($"\n--- Используется неверный междустрочный интервал: {(p.LineSpacing / 12).ToString("0.00").Replace(',', '.')} вместо {(DefaultLineSpacing / 12).ToString("0.00").Replace(',', '.')} строки ---");
                        }
                    }
                }
            }
            catch
            {
                ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                    " - предположительно используется неверный междустрочный интервал." + '\n';
                if (checkCopy)
                {
                    if (checkWarnings)
                    {
                        p.InsertText("\n--- Предположительно используется неверный междустрочный интервал. ---");
                    }
                }
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
                        if (checkCopy)
                        {
                            if (checkErrors)
                            {
                                p.InsertText("\n--- Используется выравнивание по центру. ---");
                            }
                        }
                    }
                    else if (p.Alignment.ToString() == "left")
                    {
                        ResultErrors += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            " - используется выравнивание по левому краю." + '\n';
                        if (checkCopy)
                        {
                            if (checkErrors)
                            {
                                p.InsertText("\n--- Используется выравнивание по левому краю. ---");
                            }
                        }
                    }
                    else if (p.Alignment.ToString() == "right")
                    {
                        ResultErrors += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            " - используется выравнивание по правому краю." + '\n';
                        if (checkCopy)
                        {
                            if (checkErrors)
                            {
                                p.InsertText("\n--- Используется выравнивание по правому краю. ---");
                            }
                        }
                    }
                    else if (p.Alignment.ToString() == "both")
                    {
                        ResultErrors += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            " - используется выравнивание по ширине." + '\n';
                        if (checkCopy)
                        {
                            if (checkErrors)
                            {
                                p.InsertText("\n--- Используется выравнивание по ширине. ---");
                            }
                        }
                    }
                    else
                    {
                        ResultWarnings += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                            " - предположительно используется неверное выравнивание." + '\n';
                        if (checkCopy)
                        {
                            if (checkWarnings)
                            {
                                p.InsertText("\n--- Предположительно используется неверное выравнивание. ---");
                            }
                        }
                    }
                }
            }
            catch { }
        }


        void SetDefaults(string filePath)
        {

            using (WordprocessingDocument docx = WordprocessingDocument.Open(filePath, true))
            {
                fontTable = docx.MainDocumentPart.FontTablePart.Fonts;
                //styles = docx.MainDocumentPart.StyleDefinitionsPart.Styles.ChildElements;

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
                docx.Close();
            }

        }

        bool GetTableOfContentsInfo(string filePath)
        {
            using (var doc = DocX.Load(filePath))
            {
                foreach (var p in doc.Paragraphs)
                {
                    if (p.Text.ToUpper().Trim(' ').Equals("СОДЕРЖАНИЕ") || p.Text.ToUpper().Trim(' ').Equals("ОГЛАВЛЕНИЕ"))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public void MyDocument(string filePath, bool checkFonts = true, bool checkFontSize = true,
            bool checkLineSpacing = true, bool checkAlignment = true, bool copy = false, bool errors = false, bool warnings = false, bool title = true)
        {
            try
            {
                SetDefaults(filePath);
                checkCopy = copy;
                checkErrors = errors;
                checkWarnings = warnings;

                if (File.Exists(filePath))
                {
                    if (checkCopy)
                    {
                        string tempFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) +
                            filePath.Remove(0, filePath.LastIndexOf('\\')).Replace(".docx", "_copy.docx");

                        File.Copy(filePath, tempFolder, true);
                        filePath = tempFolder;
                    }

                    bool tableOfContentsInfo = GetTableOfContentsInfo(filePath);
                    using (var doc = DocX.Load(filePath))
                    {
                        uint line = 0;
                        string text;
                        bool findTable = false;
                        foreach (var p in doc.Paragraphs)
                        {
                            line++;

                            if (title)
                            {
                                if (tableOfContentsInfo)
                                {
                                    if (p.Text.ToUpper().Trim(' ').Equals("СОДЕРЖАНИЕ") || p.Text.ToUpper().Trim(' ').Equals("ОГЛАВЛЕНИЕ"))
                                    {
                                        findTable = true;
                                        continue; // Оглавление найдено
                                    }
                                    else if (!findTable)
                                    {
                                        continue;
                                    }
                                    else if (p.Xml.ToString().Contains("hyperlink") || p.Xml.ToString().Contains("fldSimple") || p.Xml.ToString().Contains("fldChar"))
                                    {
                                        continue;
                                    }
                                    tableOfContentsInfo = false;
                                }
                            }

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
                        if (checkCopy)
                        {
                            doc.Save();
                            AssignedProcess = new Process();
                            AssignedProcess.StartInfo.FileName = filePath;
                            AssignedProcess.StartInfo.Verb = "Open";
                            AssignedProcess.StartInfo.CreateNoWindow = false;
                            AssignedProcess.Start();
                        }
                    }
                }
            }
            catch
            {
                ResultErrors = ResultWarnings = "Выбран некорректный тип файла!" + '\n' +
                    "Убедитесь, что выбран файл формата DOCX. " + '\n' +
                    "При необоходимости откройте Microsoft Word и пересохраните Ваш документ в нужном формате." + '\n' +
                    "Возможно, файл используется другим процессом (например, Microsoft Word).";
            }
        }
    }
}