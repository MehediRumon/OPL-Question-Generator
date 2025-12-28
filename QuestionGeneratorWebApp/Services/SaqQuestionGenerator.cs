using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Xml.Linq; // added for XML manipulation
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Vml;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

public static class SaqQuestionGenerator
{
    public static List<string> CreateSaqQuestion(
        string outputDirectoryPath,
        int questionGenerateNumber,
        int questionMark,
        string subjectListBangla,
        string subjectListEnglish,
        string sequenceList,
        bool multiSet = false,
        int setCount = 1
    )
    {
        string saqSamplePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "Question", "SaqSample.docx");
        if (!File.Exists(saqSamplePath))
            throw new FileNotFoundException($"SaqSample.docx not found at '{saqSamplePath}'");

        Directory.CreateDirectory(outputDirectoryPath);

        var bnSubjects = subjectListBangla.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(s => s.Trim()).ToList();
        var enSubjects = subjectListEnglish.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(s => s.Trim()).ToList();

        if (bnSubjects.Count != enSubjects.Count)
            throw new ArgumentException("subjectListBangla এবং subjectListEnglish এ উপাদানের সংখ্যা সমান থাকতে হবে।");

        var ranges = ParseRanges(sequenceList);
        var createdFiles = new List<string>();
        string batchId = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        int sets = (multiSet && setCount > 1) ? setCount : 1;

        using (var sampleDoc = WordprocessingDocument.Open(saqSamplePath, false))
        {
            var sampleMain = sampleDoc.MainDocumentPart;
            var sampleTables = sampleMain.Document.Body.Descendants<W.Table>().ToList();
            if (sampleTables.Count == 0)
                throw new InvalidOperationException("SaqSample থেকে কোনও টেবিল পাওয়া যায়নি।");

            int sampleCount = sampleTables.Count;

            if (sets == 1)
            {
                string outFileName = $"SAQ_{batchId}.docx";
                string outPath = System.IO.Path.Combine(outputDirectoryPath, outFileName);

                using (var outDoc = WordprocessingDocument.Create(outPath, WordprocessingDocumentType.Document))
                {
                    var outMain = outDoc.AddMainDocumentPart();
                    outMain.Document = new W.Document(new W.Body());
                    var body = outMain.Document.Body;

                    StripHeaderFooterDrawings(outMain);

                    for (int s = 0; s < bnSubjects.Count; s++)
                    {
                        string subjBn = bnSubjects[s];
                        string subjEn = enSubjects[s];

                        foreach (var (start, end) in ranges)
                        {
                            int rangeSize = Math.Abs(end - start) + 1;
                            int neededCount = questionGenerateNumber > 0 ? Math.Min(rangeSize, questionGenerateNumber) : rangeSize;

                            for (int qi = 0; qi < neededCount; qi++)
                            {
                                int serialNumber = start + qi;
                                int sourceIndex = (serialNumber - 1) % sampleCount;
                                var newTable = CopyTableWithImages(sampleTables[sourceIndex], sampleDoc, sampleMain, outMain);
                                EnsureTableBorders(newTable);

                                ReplaceBanglaSubjectCell(newTable, 0, 0, subjBn);
                                ReplaceCellText(newTable, 0, 1, subjEn, bangla: false);
                                ReplaceCellText(newTable, 1, 0, serialNumber.ToString(), bangla: false);
                                ReplaceCellText(newTable, 1, 1, questionMark.ToString(), bangla: false);

                                body.AppendChild(newTable);
                                body.AppendChild(new W.Paragraph(new W.Run(new W.Text(" "))));
                            }
                        }
                    }

                    SanitizeDocument(outDoc);
                    EnsureWordDefaults(outDoc);
                    EnsureSectionPrLast(outDoc);
                    ForceCompatibilityMode15(outDoc);

                    outMain.Document.Save();
                }

                createdFiles.Add(outFileName);
            }
            else
            {
                for (int setIndex = 1; setIndex <= sets; setIndex++)
                {
                    string outFileName = $"SAQ_Set{setIndex}_{batchId}.docx";
                    string outPath = System.IO.Path.Combine(outputDirectoryPath, outFileName);

                    using (var outDoc = WordprocessingDocument.Create(outPath, WordprocessingDocumentType.Document))
                    {
                        var outMain = outDoc.AddMainDocumentPart();
                        outMain.Document = new W.Document(new W.Body());
                        var body = outMain.Document.Body;

                        StripHeaderFooterDrawings(outMain);

                        for (int s = 0; s < bnSubjects.Count; s++)
                        {
                            string subjBn = bnSubjects[s];
                            string subjEn = enSubjects[s];

                            foreach (var (start, end) in ranges)
                            {
                                int rangeSize = Math.Abs(end - start) + 1;
                                int neededCount = questionGenerateNumber > 0 ? Math.Min(rangeSize, questionGenerateNumber) : rangeSize;

                                for (int qi = 0; qi < neededCount; qi++)
                                {
                                    int serialNumber = start + qi;
                                    int sourceIndex = (serialNumber - 1) % sampleCount;
                                    var newTable = CopyTableWithImages(sampleTables[sourceIndex], sampleDoc, sampleMain, outMain);
                                    EnsureTableBorders(newTable);

                                    ReplaceBanglaSubjectCell(newTable, 0, 0, subjBn);
                                    ReplaceCellText(newTable, 0, 1, subjEn, bangla: false);
                                    ReplaceCellText(newTable, 1, 0, serialNumber.ToString(), bangla: false);
                                    ReplaceCellText(newTable, 1, 1, questionMark.ToString(), bangla: false);
                                    AppendLatinLabelRunToRow(newTable, 2, $" [Set-{setIndex}] ");

                                    body.AppendChild(newTable);
                                    body.AppendChild(new W.Paragraph(new W.Run(new W.Text(" "))));
                                }
                            }
                        }

                        SanitizeDocument(outDoc);
                        EnsureWordDefaults(outDoc);
                        EnsureSectionPrLast(outDoc);
                        ForceCompatibilityMode15(outDoc);

                        outMain.Document.Save();
                    }

                    createdFiles.Add(outFileName);
                }
            }
        }

        return createdFiles;
    }

    // Force w:compat/w:compatSetting w:name="compatibilityMode" w:val="15" using raw XML
    private static void ForceCompatibilityMode15(WordprocessingDocument doc)
    {
        var settingsPart = doc.MainDocumentPart.DocumentSettingsPart ?? doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
        if (settingsPart.Settings == null)
        {
            settingsPart.Settings = new W.Settings(new W.Compatibility());
            settingsPart.Settings.Save();
        }

        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        XDocument xdoc;
        using (var s = settingsPart.GetStream())
        {
            xdoc = XDocument.Load(s);
        }

        var root = xdoc.Element(w + "settings");
        if (root == null)
        {
            root = new XElement(w + "settings");
            xdoc.Add(root);
        }
        var compat = root.Element(w + "compat");
        if (compat == null)
        {
            compat = new XElement(w + "compat");
            root.Add(compat);
        }
        // remove existing compatibilityMode setting
        compat.Elements(w + "compatSetting")
              .Where(e => (string)e.Attribute(w + "name") == "compatibilityMode")
              .Remove();
        // add mode 15
        compat.Add(new XElement(w + "compatSetting",
            new XAttribute(w + "name", "compatibilityMode"),
            new XAttribute(w + "uri", "http://schemas.microsoft.com/office/word"),
            new XAttribute(w + "val", "15")));

        using (var s = settingsPart.GetStream(FileMode.Create, FileAccess.Write))
        {
            xdoc.Save(s);
        }
    }

    private static void EnsureSectionPrLast(WordprocessingDocument doc)
    {
        var body = doc.MainDocumentPart.Document.Body;
        var sectPr = body.Elements<W.SectionProperties>().LastOrDefault();
        if (sectPr == null)
        {
            sectPr = new W.SectionProperties(new W.PageSize() { Width = 11906, Height = 16838 }, new W.PageMargin() { Top = 720, Right = 720, Bottom = 720, Left = 720 });
            body.AppendChild(sectPr);
        }
        else if (body.LastChild != sectPr)
        {
            sectPr.Remove();
            body.AppendChild(sectPr);
        }
    }

    private static void StripHeaderFooterDrawings(MainDocumentPart main)
    {
        foreach (var hp in main.HeaderParts.ToList())
        {
            var hdr = hp.Header;
            if (hdr == null) continue;
            foreach (var d in hdr.Descendants<W.Drawing>().ToList()) d.Remove();
            foreach (var v in hdr.Descendants<DocumentFormat.OpenXml.Vml.Shape>().ToList()) v.Remove();
            foreach (var v in hdr.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().ToList()) v.Remove();
            foreach (var unk in hdr.Descendants<OpenXmlUnknownElement>().ToList())
            {
                if (string.Equals(unk.LocalName, "pict", StringComparison.OrdinalIgnoreCase)) unk.Remove();
            }
            hp.Header.Save();
        }
        foreach (var fp in main.FooterParts.ToList())
        {
            var ftr = fp.Footer;
            if (ftr == null) continue;
            foreach (var d in ftr.Descendants<W.Drawing>().ToList()) d.Remove();
            foreach (var v in ftr.Descendants<DocumentFormat.OpenXml.Vml.Shape>().ToList()) v.Remove();
            foreach (var v in ftr.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().ToList()) v.Remove();
            foreach (var unk in ftr.Descendants<OpenXmlUnknownElement>().ToList())
            {
                if (string.Equals(unk.LocalName, "pict", StringComparison.OrdinalIgnoreCase)) unk.Remove();
            }
            fp.Footer.Save();
        }
    }

    private static void EnsureWordDefaults(WordprocessingDocument doc)
    {
        var body = doc.MainDocumentPart.Document.Body;
        if (body.GetFirstChild<W.SectionProperties>() == null)
        {
            body.AppendChild(new W.SectionProperties(
                new W.PageSize() { Width = 11906, Height = 16838 },
                new W.PageMargin() { Top = 720, Right = 720, Bottom = 720, Left = 720 }
            ));
        }

        if (doc.MainDocumentPart.StyleDefinitionsPart == null)
        {
            var stylesPart = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            var styles = new W.Styles(
                new W.DocDefaults(
                    new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(new W.RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Calibri" }, new W.FontSize() { Val = "22" })),
                    new W.ParagraphPropertiesDefault(new W.ParagraphProperties())
                ),
                new W.Style(new W.StyleName() { Val = "Normal" }) { Type = W.StyleValues.Paragraph, Default = true, StyleId = "Normal" }
            );
            stylesPart.Styles = styles;
            stylesPart.Styles.Save();
        }

        if (doc.MainDocumentPart.DocumentSettingsPart == null)
        {
            var settingsPart = doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new W.Settings(new W.Compatibility());
            settingsPart.Settings.Save();
        }

        if (doc.ExtendedFilePropertiesPart == null)
        {
            var app = doc.AddNewPart<ExtendedFilePropertiesPart>();
            app.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties(
                new DocumentFormat.OpenXml.ExtendedProperties.Application("Microsoft Office Word")
            );
            app.Properties.Save();
        }

        if (!doc.PackageProperties.Created.HasValue)
        {
            doc.PackageProperties.Created = DateTime.UtcNow;
        }
    }

    private static void SanitizeDocument(WordprocessingDocument doc)
    {
        foreach (var part in doc.Parts.Select(p => p.OpenXmlPart).Concat(new[] { (OpenXmlPart)doc.MainDocumentPart }))
        {
            if (part is MainDocumentPart mdp)
            {
                foreach (var rel in mdp.HyperlinkRelationships.ToList()) mdp.DeleteReferenceRelationship(rel);
                foreach (var r in mdp.ExternalRelationships.ToList()) mdp.DeleteExternalRelationship(r);
            }

            var blips = part.RootElement?.Descendants<A.Blip>().ToList();
            if (blips != null)
            {
                foreach (var b in blips)
                {
                    b.Link = null;
                }
            }
        }
    }

    private static void ReplaceBanglaSubjectCell(W.Table table, int rowIndex, int colIndex, string text)
    {
        var rows = table.Elements<W.TableRow>().ToList();
        if (rowIndex < 0 || rowIndex >= rows.Count) return;
        var cells = rows[rowIndex].Elements<W.TableCell>().ToList();
        if (colIndex < 0 || colIndex >= cells.Count) return;
        var cell = cells[colIndex];

        cell.RemoveAllChildren<W.Paragraph>();
        var para = new W.Paragraph();
        var run = new W.Run();
        var rp = new W.RunProperties();
        rp.AppendChild(new W.Languages { Val = "bn-BD", EastAsia = "bn-BD", Bidi = "bn-BD" });
        rp.AppendChild(new W.RunFonts { Ascii = "SutonnyMJ", HighAnsi = "SutonnyMJ", EastAsia = "SutonnyMJ", ComplexScript = "SutonnyMJ" });
        rp.AppendChild(new W.RightToLeftText());
        run.AppendChild(rp);
        run.AppendChild(new W.Text(text ?? string.Empty));
        para.AppendChild(run);
        cell.AppendChild(para);
    }

    private static ImagePart TryResolveImagePart(WordprocessingDocument sourceDoc, OpenXmlPart containerPart, string relId)
    {
        if (string.IsNullOrEmpty(relId)) return null;
        try { var part = containerPart.GetPartById(relId) as ImagePart; if (part != null) return part; } catch {}
        try { var part = sourceDoc.MainDocumentPart.GetPartById(relId) as ImagePart; if (part != null) return part; } catch {}
        var relatedParts = new List<OpenXmlPart>();
        relatedParts.AddRange(sourceDoc.MainDocumentPart.HeaderParts);
        relatedParts.AddRange(sourceDoc.MainDocumentPart.FooterParts);
        relatedParts.AddRange(sourceDoc.MainDocumentPart.ChartParts);
        relatedParts.AddRange(sourceDoc.MainDocumentPart.ImageParts);
        foreach (var rp in relatedParts)
        {
            try { var p = rp.GetPartById(relId) as ImagePart; if (p != null) return p; } catch {}
        }
        foreach (var p in sourceDoc.Parts.Select(pp => pp.OpenXmlPart))
        {
            if (p is ImagePart ip)
            {
                try { var id = sourceDoc.MainDocumentPart.GetIdOfPart(ip); if (id == relId) return ip; } catch {}
            }
        }
        return null;
    }

    private static string NormalizeImageContentType(string contentType)
    {
        if (string.IsNullOrEmpty(contentType)) return "image/png";
        var ct = contentType.ToLowerInvariant();
        if (ct == "image/png" || ct == "image/jpeg") return ct;
        return "image/png";
    }

    private static void CopyImageDataToPart(ImagePart sourceImgPart, ImagePart targetImgPart)
    {
        var ct = sourceImgPart.ContentType.ToLowerInvariant();
        if (ct == "image/png" || ct == "image/jpeg")
        {
            using (var s = sourceImgPart.GetStream())
            using (var t = targetImgPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                s.CopyTo(t);
            }
        }
        else
        {
            using (var s = sourceImgPart.GetStream())
            using (var img = Image.FromStream(s, useEmbeddedColorManagement: false, validateImageData: false))
            using (var t = targetImgPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                img.Save(t, System.Drawing.Imaging.ImageFormat.Png);
            }
        }
    }

    private static W.Table CopyTableWithImages(W.Table sourceTable, WordprocessingDocument sourceDoc, MainDocumentPart sourceMain, MainDocumentPart targetMain)
    {
        var cloned = (W.Table)sourceTable.CloneNode(true);

        foreach (var vShape in cloned.Descendants<DocumentFormat.OpenXml.Vml.Shape>().ToList()) vShape.Remove();
        foreach (var vImg in cloned.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().ToList()) vImg.Remove();
        foreach (var unk in cloned.Descendants<OpenXmlUnknownElement>().ToList())
        {
            if (string.Equals(unk.LocalName, "pict", StringComparison.OrdinalIgnoreCase)) unk.Remove();
        }

        foreach (var drawing in cloned.Descendants<W.Drawing>())
        {
            var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
            if (blip != null)
            {
                var relId = blip.Embed?.Value ?? blip.Link?.Value;
                var imgPart = TryResolveImagePart(sourceDoc, sourceMain, relId);
                if (imgPart != null)
                {
                    var newContentType = NormalizeImageContentType(imgPart.ContentType);
                    var newImgPart = targetMain.AddImagePart(newContentType);
                    CopyImageDataToPart(imgPart, newImgPart);
                    var newRelId = targetMain.GetIdOfPart(newImgPart);
                    blip.Embed = newRelId;
                    blip.Link = null;
                }
                else
                {
                    blip.Remove();
                }
            }
        }
        return cloned;
    }

    private static void EnsureTableBorders(W.Table table)
    {
        var props = table.GetFirstChild<W.TableProperties>();
        if (props == null)
        {
            props = new W.TableProperties();
            table.PrependChild(props);
        }
        var borders = props.GetFirstChild<W.TableBorders>();
        if (borders == null)
        {
            borders = new W.TableBorders();
            props.AppendChild(borders);
        }
        borders.TopBorder = new W.TopBorder { Val = W.BorderValues.Single, Size = 8 };
        borders.BottomBorder = new W.BottomBorder { Val = W.BorderValues.Single, Size = 8 };
        borders.LeftBorder = new W.LeftBorder { Val = W.BorderValues.Single, Size = 8 };
        borders.RightBorder = new W.RightBorder { Val = W.BorderValues.Single, Size = 8 };
        borders.InsideHorizontalBorder = new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = 8 };
        borders.InsideVerticalBorder = new W.InsideVerticalBorder { Val = W.BorderValues.Single, Size = 8 };
    }

    private static void ReplaceCellText(W.Table table, int rowIndex, int colIndex, string text, bool bangla)
    {
        var rows = table.Elements<W.TableRow>().ToList();
        if (rowIndex < 0 || rowIndex >= rows.Count) return;
        var cells = rows[rowIndex].Elements<W.TableCell>().ToList();
        if (colIndex < 0 || colIndex >= cells.Count) return;
        var cell = cells[colIndex];

        cell.RemoveAllChildren<W.Paragraph>();
        var para = new W.Paragraph();
        var run = new W.Run();
        var rp = new W.RunProperties();
        if (bangla)
        {
            rp.AppendChild(new W.Languages { Val = "bn-BD", EastAsia = "bn-BD", Bidi = "bn-BD" });
            rp.AppendChild(new W.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Calibri", ComplexScript = "SutonnyMJ" });
            rp.AppendChild(new W.RightToLeftText());
        }
        else
        {
            rp.AppendChild(new W.Languages { Val = "en-US", EastAsia = "en-US", Bidi = "en-US" });
            rp.AppendChild(new W.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Calibri" });
        }
        run.AppendChild(rp);
        run.AppendChild(new W.Text(text));
        para.AppendChild(run);
        cell.AppendChild(para);
    }

    private static void OverwriteCellText(W.Table table, int rowIndex, int colIndex, string text, bool bangla)
    {
        var rows = table.Elements<W.TableRow>().ToList();
        if (rowIndex < 0 || rowIndex >= rows.Count) return;
        var cells = rows[rowIndex].Elements<W.TableCell>().ToList();
        if (colIndex < 0 || colIndex >= cells.Count) return;
        var cell = cells[colIndex];

        var para = cell.Descendants<W.Paragraph>().FirstOrDefault();
        if (para == null)
        {
            para = new W.Paragraph();
            cell.AppendChild(para);
        }
        var run = new W.Run();
        var rp = new W.RunProperties();
        if (bangla)
        {
            rp.AppendChild(new W.Languages { Val = "bn-BD", EastAsia = "bn-BD", Bidi = "bn-BD" });
            rp.AppendChild(new W.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Calibri", ComplexScript = "SutonnyMJ" });
            rp.AppendChild(new W.RightToLeftText());
        }
        else
        {
            rp.AppendChild(new W.Languages { Val = "en-US", EastAsia = "en-US", Bidi = "en-US" });
            rp.AppendChild(new W.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Calibri" });
        }
        run.AppendChild(rp);
        run.AppendChild(new W.Text(text));
        para.AppendChild(run);
    }

    private static void AppendLatinLabelRunToRow(W.Table table, int rowIndex, string labelText)
    {
        var rows = table.Elements<W.TableRow>().ToList();
        if (rowIndex < 0 || rowIndex >= rows.Count) return;
        var row = rows[rowIndex];
        var cells = row.Elements<W.TableCell>().ToList();
        foreach (var cell in cells.Take(2))
        {
            var para = cell.Descendants<W.Paragraph>().LastOrDefault();
            if (para == null)
            {
                para = new W.Paragraph();
                cell.AppendChild(para);
            }
            var run = new W.Run();
            var rp = new W.RunProperties();
            rp.AppendChild(new W.Languages { Val = "en-US", EastAsia = "en-US", Bidi = "en-US" });
            rp.AppendChild(new W.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Calibri" });
            run.AppendChild(rp);
            run.AppendChild(new W.Text(labelText));
            para.AppendChild(run);
        }
    }

    private static List<(int start, int end)> ParseRanges(string sequenceList)
    {
        var outL = new List<(int, int)>();
        if (string.IsNullOrWhiteSpace(sequenceList)) return outL;
        var tokens = sequenceList.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var t in tokens)
        {
            var s = t.Trim();
            var parts = s.Split('-');
            if (parts.Length == 2 && int.TryParse(parts[0], out int a) && int.TryParse(parts[1], out int b))
            {
                outL.Add((a, b));
            }
            else
            {
                throw new ArgumentException($"Invalid range token: {s}. Expect format like 1-10.");
            }
        }
        return outL;
    }
}
