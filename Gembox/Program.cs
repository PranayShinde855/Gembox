using System;
using System.IO;
using System.Collections.Generic;
using GemBox.Document;

class Program
{
    static void Main()
    {
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var document = new DocumentModel();

        // -----------------------------
        // STORE BOOKMARKS
        // -----------------------------
        var tocItems = new List<(string Title, string Bookmark, int Level)>();

        // -----------------------------
        // CREATE TOC SECTION
        // -----------------------------
        var tocSection = new Section(document);
        document.Sections.Add(tocSection);

        tocSection.Blocks.Add(new Paragraph(document, "Table of Contents")
        {
            ParagraphFormat = { Alignment = HorizontalAlignment.Center }
        });

        // -----------------------------
        // CREATE CONTENT SECTION
        // -----------------------------
        var contentSection = new Section(document);
        document.Sections.Add(contentSection);

        int h1 = 1;

        for (int i = 1; i <= 3; i++)
        {
            string bm1 = $"H1_{i}";
            string title1 = $"{h1}. Heading {i}";

            AddHeading(document, contentSection, title1, bm1, 1);
            tocItems.Add((title1, bm1, 1));

            int h2 = 1;

            for (int j = 1; j <= 2; j++)
            {
                string bm2 = $"H2_{i}_{j}";
                string title2 = $"{h1}.{h2}. Sub Heading {j}";

                AddHeading(document, contentSection, title2, bm2, 2);
                tocItems.Add((title2, bm2, 2));

                int h3 = 1;

                for (int k = 1; k <= 2; k++)
                {
                    string bm3 = $"H3_{i}_{j}_{k}";
                    string title3 = $"{h1}.{h2}.{h3}. Child {k}";

                    AddHeading(document, contentSection, title3, bm3, 3);
                    tocItems.Add((title3, bm3, 3));

                    h3++;
                }

                h2++;
            }

            h1++;
        }

        // -----------------------------
        // BUILD CLICKABLE TOC
        // -----------------------------
        ////foreach (var item in tocItems)
        ////{
        ////    // indentation
        ////    string indent = new string('\t', item.Level - 1);

        ////    // create field
        ////    Field field = new Field(document, FieldType.Hyperlink);

        ////    // IMPORTANT: use Code, not Instruction
        ////    field.Code = "HYPERLINK \\l \"" + item.Bookmark + "\"";

        ////    // create styled text
        ////    Run run = new Run(document, indent + item.Title);
        ////    run.CharacterFormat.FontColor = Color.Blue;
        ////    run.CharacterFormat.UnderlineStyle = UnderlineType.Single;

        ////    // attach text to field
        ////    field.ResultInlines.Add(run);

        ////    // create paragraph
        ////    Paragraph p = new Paragraph(document, field);

        ////    // add to TOC
        ////    tocSection.Blocks.Add(p);
        ////}
        ///
        foreach (var item in tocItems)
        {
            Paragraph p = new Paragraph(document);

            // indentation using tabs (safe for v35)
            string indent = new string('\t', item.Level - 1);

            // create hyperlink (this works in v35)
            Hyperlink link = new Hyperlink(document, indent + item.Title, item.Bookmark);

            p.Inlines.Add(link);

            tocSection.Blocks.Add(p);
        }

        // -----------------------------
        // SAVE PDF
        // -----------------------------
        document.Save("Final_TOC.pdf");
    }

    // -----------------------------
    // HELPER METHOD
    // -----------------------------
    static void AddHeading(DocumentModel document, Section section, string text, string bookmark, int level)
    {
        var para = new Paragraph(document);

        para.Inlines.Add(new BookmarkStart(document, bookmark));
        para.Inlines.Add(new Run(document, text)
        {
            CharacterFormat = { Bold = true, Size = 14 }
        });
        para.Inlines.Add(new BookmarkEnd(document, bookmark));

        section.Blocks.Add(para);

        section.Blocks.Add(new Paragraph(document,
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit.\n\n"));
    }
}




















//////using System;
//////using System.IO;
//////using System.Linq;
//////using System.Collections.Generic;
//////using GemBox.Document;
//////using GemBox.Pdf;
//////using GemBox.Pdf.Content;
//////using GemBox.Pdf.Annotations;
//////using ComponentInfo = GemBox.Document.ComponentInfo;
//////using SaveOptions = GemBox.Document.SaveOptions;
//////using GemBox.Pdf;
//////using GemBox.Pdf.Content;
//////using GemBox.Pdf.Annotations;
//////using GemBox.Pdf.Actions;

//////class Program
//////{
//////    static void Main()
//////    {
//////        // License
//////        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
//////        GemBox.Pdf.ComponentInfo.SetLicense("FREE-LIMITED-KEY");

//////        // -----------------------------
//////        // TOC DATA (Replace with your AgendaDetails)
//////        // -----------------------------
//////        var tocItems = new List<TocItem>
//////        {
//////            new TocItem { Title = "1. Introduction", Level = 1, TargetPageIndex = 0 },

//////            new TocItem { Title = "2. Agenda", Level = 1, TargetPageIndex = 1 },
//////            new TocItem { Title = "2.1 Overview", Level = 2, TargetPageIndex = 2 },
//////            new TocItem { Title = "2.2 Scope", Level = 2, TargetPageIndex = 3 },
//////            new TocItem { Title = "2.2.1 Internal Scope", Level = 3, TargetPageIndex = 4 },
//////            new TocItem { Title = "2.2.2 External Scope", Level = 3, TargetPageIndex = 5 },

//////            new TocItem { Title = "3. Financials", Level = 1, TargetPageIndex = 6 }
//////        };

//////        // -----------------------------
//////        // 1. CREATE TOC PDF
//////        // -----------------------------
//////        MemoryStream tocStream = new MemoryStream();
//////        CreateTocPdf(tocItems, tocStream);

//////        // -----------------------------
//////        // 2. LOAD CONTENT PDF
//////        // -----------------------------
//////        FileStream contentStream = File.OpenRead("Content.pdf");

//////        // -----------------------------
//////        // 3. MERGE PDFs
//////        // -----------------------------
//////        PdfDocument finalPdf = MergePdfs(tocStream, contentStream);

//////        // -----------------------------
//////        // 4. ADD CLICKABLE LINKS
//////        // -----------------------------
//////        AddTocLinks(finalPdf, tocItems, 1);

//////        // -----------------------------
//////        // 5. SAVE
//////        // -----------------------------
//////        finalPdf.Save("Final_Output.pdf");

//////        contentStream.Close();
//////        tocStream.Close();
//////    }

//////    // -----------------------------
//////    // MODEL
//////    // -----------------------------
//////    public class TocItem
//////    {
//////        public string Title { get; set; }
//////        public int Level { get; set; }
//////        public int TargetPageIndex { get; set; }
//////    }

//////    // -----------------------------
//////    // CREATE TOC PDF (GemBox.Document)
//////    // -----------------------------
//////    static void CreateTocPdf(List<TocItem> items, Stream output)
//////    {
//////        DocumentModel document = new DocumentModel();
//////        Section section = new Section(document);
//////        document.Sections.Add(section);

//////        Paragraph title = new Paragraph(document, "Table of Contents");
//////        title.ParagraphFormat.Alignment = HorizontalAlignment.Center;
//////        section.Blocks.Add(title);
//////        foreach (var item in items)
//////        {
//////            Paragraph p = new Paragraph(document);

//////            // Indentation using spaces (version-safe)
//////            string indentText = new string(' ', (item.Level - 1) * 4);

//////            Run run = new Run(document, indentText + item.Title);
//////            run.CharacterFormat.FontColor = Color.Blue;
//////            run.CharacterFormat.UnderlineStyle = UnderlineType.Single;

//////            p.Inlines.Add(run);

//////            section.Blocks.Add(p);
//////        }

//////        document.Save(output, SaveOptions.PdfDefault);
//////        output.Position = 0;
//////    }

//////    // -----------------------------
//////    // MERGE PDFs (GemBox.Pdf)
//////    // -----------------------------
//////    static PdfDocument MergePdfs(Stream tocStream, Stream contentStream)
//////    {
//////        PdfDocument finalPdf = new PdfDocument();

//////        PdfDocument tocPdf = PdfDocument.Load(tocStream);
//////        PdfDocument contentPdf = PdfDocument.Load(contentStream);

//////        // Add TOC pages
//////        for (int i = 0; i < tocPdf.Pages.Count; i++)
//////            finalPdf.Pages.AddClone(tocPdf.Pages[i]);

//////        // Add content pages
//////        for (int i = 0; i < contentPdf.Pages.Count; i++)
//////            finalPdf.Pages.AddClone(contentPdf.Pages[i]);

//////        return finalPdf;
//////    }

//////    // -----------------------------
//////    // ADD CLICKABLE LINKS
//////    // -----------------------------
//////    static void AddTocLinks(PdfDocument pdf, List<TocItem> items, int tocPageCount)
//////    {
//////        PdfPage tocPage = pdf.Pages[0];

//////        int contentStartIndex = tocPageCount;

//////        double startY = 750;   // adjust if needed
//////        double lineHeight = 18;

//////        for (int i = 0; i < items.Count; i++)
//////        {
//////            TocItem item = items[i];

//////            int targetIndex = contentStartIndex + item.TargetPageIndex;

//////            if (targetIndex >= pdf.Pages.Count)
//////                continue;

//////            PdfPage targetPage = pdf.Pages[targetIndex];

//////            // Correct action class
//////            PdfGoToPageAction action = new PdfGoToPageAction(targetPage);

//////            // Clickable area
//////            PdfRectangle rect = new PdfRectangle(x, y, 400, 15);

//////            // Link annotation
//////            PdfLinkAnnotation link = new PdfLinkAnnotation(tocPage, rect, action);

//////            // Add to page
//////            tocPage.Annotations.Add(link);

//////            // X position based on level (indent)
//////            double x = 50 + (item.Level - 1) * 20;
//////            double y = startY - (i * lineHeight);

//////            PdfRectangle rect = new PdfRectangle(x, y, 400, 15);

//////            PdfLinkAnnotation link = new PdfLinkAnnotation(tocPage, rect, action);

//////            tocPage.Annotations.Add(link);
//////        }


//////    }
//////}




public class PdfBuilder
{
    public MemoryStream Build(PdfTableOfContentDto requestDto)
    {
        var finalDocument = new PdfDocument();
        finalDocument.ViewerPreferences.NonFullScreenPageMode = PdfPageMode.UseOutlines;

        var tocEntries = new List<(string Title, int PagesCount)>();
        int totalContentPages = 0;

        // -----------------------------
        // 1. MERGE CONTENT PDFs
        // -----------------------------
        int agendaIndex = 1;

        foreach (var agenda in requestDto.AgendaDetails)
        {
            AddDocumentContent(finalDocument, agenda.Stream,
                $"{agendaIndex}. {agenda.LabelText}",
                tocEntries, ref totalContentPages, agenda);

            int subIndex = 1;

            foreach (var sub in agenda.SubAgenda)
            {
                AddDocumentContent(finalDocument, sub.Stream,
                    $"{agendaIndex}.{subIndex}. {sub.LabelText}",
                    tocEntries, ref totalContentPages, sub);

                int childIndex = 1;

                foreach (var child in sub.SubChildAgenda)
                {
                    AddDocumentContent(finalDocument, child.Stream,
                        $"{agendaIndex}.{subIndex}.{childIndex++}. {child.LabelText}",
                        tocEntries, ref totalContentPages, child);
                }

                subIndex++;
            }

            agendaIndex++;
        }

        // -----------------------------
        // 2. CREATE TOC PDF
        // -----------------------------
        int headerPages, footerPages;

        var tocStream = CreateTocPdf(
            requestDto,
            requestDto.Header,
            requestDto.Footer,
            out headerPages,
            out footerPages);

        var tocDocument = PdfDocument.Load(tocStream);
        int tocPageCount = tocDocument.Pages.Count;

        // Insert TOC at beginning
        finalDocument.Pages.Kids.InsertClone(0, tocDocument.Pages);

        // -----------------------------
        // 3. BUILD PAGE MAP
        // -----------------------------
        var pageMap = new Dictionary<string, int>();
        int currentPage = tocPageCount;

        foreach (var entry in tocEntries)
        {
            pageMap[Normalize(entry.Title)] = currentPage;
            currentPage += entry.PagesCount;
        }

        // -----------------------------
        // 4. FIX OUTLINES
        // -----------------------------
        FixOutlines(finalDocument.Outlines, pageMap, finalDocument);

        // -----------------------------
        // 5. SAVE FINAL PDF
        // -----------------------------
        var resultStream = new MemoryStream();
        finalDocument.Save(resultStream, SaveOptions.Pdf);

        return resultStream;
    }

    // =========================================================
    // CREATE TOC USING DOCUMENT MODEL
    // =========================================================
    private static MemoryStream CreateTocPdf(
        PdfTableOfContentDto tocEntries,
        string headerString,
        string footerString,
        out int headerPageCount,
        out int footerPageCount)
    {
        var document = new DocumentModel();
        var section = new Section(document);
        document.Sections.Add(section);

        headerPageCount = 0;
        footerPageCount = 0;

        // HEADER
        if (!string.IsNullOrEmpty(headerString))
        {
            section.Blocks.Add(new Paragraph(document, "@header@"));

            var replace = section.Content.Find("@header@").First();
            replace.Start.LoadText(headerString, new HtmlLoadOptions()
            {
                InheritCharacterFormat = true
            });

            document.Content.Replace("@header@", "");
        }

        // TOC FIELD
        var toc = new TableOfEntries(document, FieldType.TOC);
        section.Blocks.Add(toc);

        // STYLES
        var h1 = (ParagraphStyle)document.Styles.GetOrAdd(StyleTemplateType.Heading1);
        var h2 = (ParagraphStyle)document.Styles.GetOrAdd(StyleTemplateType.Heading2);
        var h3 = (ParagraphStyle)document.Styles.GetOrAdd(StyleTemplateType.Heading3);

        h1.ParagraphFormat.PageBreakBefore = true;
        h2.ParagraphFormat.PageBreakBefore = true;
        h3.ParagraphFormat.PageBreakBefore = true;

        int i = 1;

        foreach (var agenda in tocEntries.AgendaDetails)
        {
            section.Blocks.Add(new Paragraph(document,
                $"{i}. {agenda.LabelText}")
            { ParagraphFormat = { Style = h1 } });

            int j = 1;

            foreach (var sub in agenda.SubAgenda)
            {
                section.Blocks.Add(new Paragraph(document,
                    $"{i}.{j}. {sub.LabelText}")
                { ParagraphFormat = { Style = h2 } });

                int k = 1;

                foreach (var child in sub.SubChildAgenda)
                {
                    section.Blocks.Add(new Paragraph(document,
                        $"{i}.{j}.{k++}. {child.LabelText}")
                    { ParagraphFormat = { Style = h3 } });
                }

                j++;
            }

            i++;
        }

        document.CalculateListItems();
        toc.Update();

        var options = new PdfSaveOptions
        {
            BookmarksCreateOptions = PdfBookmarksCreateOptions.UsingHeadings
        };

        var stream = new MemoryStream();
        document.Save(stream, options);

        return stream;
    }

    // =========================================================
    // ADD PDF CONTENT
    // =========================================================
    private void AddDocumentContent(
        PdfDocument mainDocument,
        Stream sourceStream,
        string title,
        List<(string Title, int PagesCount)> entriesList,
        ref int totalCount,
        dynamic agendaItem)
    {
        if (sourceStream == null)
            return;

        var sourceDoc = PdfDocument.Load(sourceStream);

        mainDocument.Pages.Kids.AddClone(sourceDoc.Pages);

        int pages = sourceDoc.Pages.Count;

        agendaItem.PageCount = pages;
        totalCount += pages;

        entriesList.Add((title, pages));
    }

    // =========================================================
    // FIX OUTLINES (RECURSIVE)
    // =========================================================
    private void FixOutlines(
        PdfOutlineCollection outlines,
        Dictionary<string, int> pageMap,
        PdfDocument document)
    {
        foreach (var outline in outlines)
        {
            var key = Normalize(outline.Title);

            if (pageMap.TryGetValue(key, out int pageIndex))
            {
                outline.SetDestination(
                    document.Pages[pageIndex],
                    PdfDestinationViewType.FitPage);
            }

            if (outline.Outlines.Count > 0)
            {
                FixOutlines(outline.Outlines, pageMap, document);
            }
        }
    }

    // =========================================================
    // NORMALIZE TITLE
    // =========================================================
    private static string Normalize(string text)
    {
        return text?.Trim().ToLower() ?? "";
    }
}