using System;

public class Class1
{
	public Class1()
	{
	}

    private string Document(string filePath)
    {
        string result = string.Empty;
        if (File.Exists(filePath))
        {
            using (WordprocessingDocument docx = WordprocessingDocument.Open(filePath, true))
            {
                DocDefaults defaults = docx.MainDocumentPart.StyleDefinitionsPart.Styles.DocDefaults;
                RunFonts runFont = defaults.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts;
                int DefaultFontSize;
                try
                {
                    DefaultFontSize = Convert.ToInt32(defaults.RunPropertiesDefault.RunPropertiesBaseStyle.FontSize.Val.Value);
                }
                catch
                {
                    DefaultFontSize = 0;
                }

                Fonts fontTable = docx.MainDocumentPart.FontTablePart.Fonts;
                docx.Close();

                using (var doc = DocX.Load(filePath))
                {
                    uint line = 0;
                    string text;
                    HashSet<string> usingFonts = new HashSet<string>();
                    HashSet<double?> usingSize = new HashSet<double?>();

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
                                        result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                    " - используется неверный шрифт: " + magic.formatting.FontFamily.Name + '\n';
                                    }
                                }
                                else if (magic.formatting.FontFamily == null & usingFonts.Contains("null") == false)
                                {
                                    usingFonts.Add("null");

                                    result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                               " - предположительно используется неверный шрифт. Возможные варианты:" + '\n';
                                    foreach (var f in fontTable.ChildElements)
                                    {
                                        Font font = f as Font;
                                        if (!font.Name.ToString().Contains("Times New Roman") & !font.Name.ToString().Contains("serif"))
                                            result += "* " + font.Name + '\n';
                                    }
                                }
                            }
                            usingFonts.Clear();
                        }
                        catch
                        {
                            result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                " - предположительно используется неверный шрифт. Возможные варианты:" + '\n';
                            foreach (var f in fontTable.ChildElements)
                            {
                                Font font = f as Font;
                                if (!font.Name.ToString().Contains("Times New Roman") & !font.Name.ToString().Contains("serif"))
                                    result += "* " + font.Name + '\n';
                            }
                        }

                        //получение кегля, работает плюс-минус нормально
                        try
                        {
                            foreach (var magic in p.MagicText)
                            {
                                if (magic.formatting.Size == null & magic.formatting.Size == 0 & usingSize.Contains(0) == false)
                                {
                                    usingSize.Add(0);
                                    if (DefaultFontSize != 0)
                                    {
                                        result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                        " - предположительно используется неверный кегль. Возможно: {DefaultFontSize/2}" + '\n';
                                    }
                                    else
                                    {
                                        result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                        " - предположительно используется неверный кегль. " + '\n';
                                    }
                                }
                                else if (magic.formatting.Size != 14 & usingSize.Contains(magic.formatting.Size) == false)
                                {
                                    usingSize.Add(magic.formatting.Size);
                                    result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                        " - используется неверный кегль: " + magic.formatting.Size + '\n';
                                }
                            }
                            usingSize.Clear();
                        }
                        catch
                        {
                            result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                    " - предположительно используется неверный кегль. " + '\n';
                        }

                        //получение междустрочного интервала, в целом работает корректно
                        try
                        {
                            if (p.LineSpacing != 18)
                            {
                                result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                    " - используется неверный междустрочный интервал: " +
                                    (p.LineSpacing / 12).ToString("0.00").Replace(',', '.') + " вместо 1.5 строки" + '\n';
                            }
                        }
                        catch (NullReferenceException)
                        {
                            result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                " - предположительно используется неверный междустрочный интервал." + '\n';
                        }
                        catch (InvalidOperationException)
                        {
                            result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                " - предположительно используется неверный междустрочный интервал." + '\n';
                        }

                        //получение информации о выравнивании, работает вроде как корректно 
                        try
                        {
                            if (p.Alignment.ToString() != "both")
                            {
                                if (p.Alignment.ToString() == "center")
                                {
                                    result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                        " - используется выравнивание по центру вместо выравнивания по ширине." + '\n';
                                }
                                else if (p.Alignment.ToString() == "left")
                                {
                                    result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                        " - используется выравнивание по левому краю вместо выравнивания по ширине." + '\n';
                                }
                                else if (p.Alignment.ToString() == "right")
                                {
                                    result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                        " - используется выравнивание по правому краю вместо выравнивания по ширине." + '\n';
                                }
                                else
                                {
                                    result += "Строка " + line + " (начинается со слов \"" + text + "\" )" +
                                        " - предположительно используется неверное выравнивание по ширине" + '\n';
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
        }

        return result;
    }
}