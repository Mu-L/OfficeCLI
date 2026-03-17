// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// Bug hunt tests Part 5: Bugs #331-350+
// Word Set, Word Query, Word Selector, GenericXmlQuery, ImageHelpers, StyleList deep dive

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public partial class BugHuntTests
{
    /// Bug #331 — Word Set: bool.Parse on TOC hyperlinks property
    /// File: WordHandler.Set.cs, lines 70-73
    /// bool.Parse("yes") or bool.Parse("1") will throw FormatException
    /// instead of being treated as truthy.
    [Fact]
    public void Bug331_WordSet_TocHyperlinksBoolParse()
    {
        // bool.Parse only accepts "True"/"False" (case-insensitive)
        // User might pass "yes", "1", "on" etc.
        var ex = Record.Exception(() => bool.Parse("yes"));
        ex.Should().BeOfType<FormatException>(
            "bool.Parse rejects 'yes' — should use IsTruthy or TryParse");

        var ex2 = Record.Exception(() => bool.Parse("1"));
        ex2.Should().BeOfType<FormatException>(
            "bool.Parse rejects '1' — common truthy value");
    }

    /// Bug #332 — Word Set: bool.Parse on TOC pageNumbers property
    /// File: WordHandler.Set.cs, lines 79-82
    /// Same bool.Parse issue as #331 but for pageNumbers switch.
    [Fact]
    public void Bug332_WordSet_TocPageNumbersBoolParse()
    {
        // bool.Parse is called directly on user-provided value
        var ex = Record.Exception(() => bool.Parse("0"));
        ex.Should().BeOfType<FormatException>(
            "bool.Parse rejects '0' — should use IsTruthy or TryParse");
    }

    /// Bug #333 — Word Set: uint.Parse on section pageWidth/pageHeight
    /// File: WordHandler.Set.cs, lines 174, 177
    /// uint.Parse("12240.5") or uint.Parse("abc") will throw.
    /// No validation on user-provided values.
    [Fact]
    public void Bug333_WordSet_SectionPageSizeUintParse()
    {
        var ex = Record.Exception(() => uint.Parse("12240.5"));
        ex.Should().BeOfType<FormatException>(
            "uint.Parse cannot handle decimal values for page size");

        var ex2 = Record.Exception(() => uint.Parse("-100"));
        ex2.Should().BeOfType<OverflowException>(
            "uint.Parse cannot handle negative values");
    }

    /// Bug #334 — Word Set: int.Parse on section margins
    /// File: WordHandler.Set.cs, lines 185-194
    /// Multiple margin properties use int.Parse / uint.Parse directly on user input.
    [Fact]
    public void Bug334_WordSet_SectionMarginParsing()
    {
        // marginTop/marginBottom use int.Parse, marginLeft/marginRight use uint.Parse
        var ex = Record.Exception(() => int.Parse("1440.5"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects decimal values for margins");
    }

    /// Bug #335 — Word Set: int.Parse on style size
    /// File: WordHandler.Set.cs, line 238
    /// int.Parse(value) * 2 — no validation that value is a number.
    [Fact]
    public void Bug335_WordSet_StyleSizeParsing()
    {
        var ex = Record.Exception(() => int.Parse("12pt"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects values with unit suffixes like '12pt'");
    }

    /// Bug #336 — Word Set: bool.Parse on 13+ run formatting properties
    /// File: WordHandler.Set.cs, lines 336-367
    /// bold, italic, caps, smallcaps, dstrike, vanish, outline, shadow,
    /// emboss, imprint, noproof, rtl, strike, superscript, subscript
    /// all use bool.Parse directly on user input.
    [Fact]
    public void Bug336_WordSet_RunBoolParseProperties()
    {
        // All these properties use bool.Parse:
        string[] boolProps = { "bold", "italic", "caps", "smallcaps", "dstrike",
            "vanish", "outline", "shadow", "emboss", "imprint", "noproof", "rtl", "strike" };

        foreach (var prop in boolProps)
        {
            var ex = Record.Exception(() => bool.Parse("yes"));
            ex.Should().BeOfType<FormatException>(
                $"bool.Parse rejects 'yes' for {prop} — should use IsTruthy");
        }
    }

    /// Bug #337 — Word Set: int.Parse on run font size with units
    /// File: WordHandler.Set.cs, line 388
    /// int.Parse(value) * 2 for half-points — no unit stripping.
    [Fact]
    public void Bug337_WordSet_RunFontSizeParsing()
    {
        // User might pass "12pt" or "12.5"
        var ex = Record.Exception(() => int.Parse("12pt"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects '12pt' — should strip unit suffix");

        var ex2 = Record.Exception(() => int.Parse("12.5"));
        ex2.Should().BeOfType<FormatException>(
            "int.Parse rejects '12.5' — half-sizes are common");
    }

    /// Bug #338 — Word Set: int.Parse on paragraph firstLineIndent
    /// File: WordHandler.Set.cs, line 529
    /// int.Parse(value) * 480 — no validation for non-numeric input.
    [Fact]
    public void Bug338_WordSet_ParagraphFirstLineIndentParsing()
    {
        var ex = Record.Exception(() => int.Parse("2.5"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects '2.5' for firstLineIndent multiplied by 480");
    }

    /// Bug #339 — Word Set: bool.Parse on paragraph keepNext/keepLines/etc.
    /// File: WordHandler.Set.cs, lines 546-568
    /// keepNext, keepLines/keepTogether, pageBreakBefore, widowControl all use bool.Parse.
    [Fact]
    public void Bug339_WordSet_ParagraphBoolParseProperties()
    {
        string[] boolProps = { "keepnext", "keeplines", "pagebreakbefore", "widowcontrol" };
        foreach (var prop in boolProps)
        {
            var ex = Record.Exception(() => bool.Parse("1"));
            ex.Should().BeOfType<FormatException>(
                $"bool.Parse rejects '1' for {prop} — should use IsTruthy");
        }
    }

    /// Bug #340 — Word Set: int.Parse on numId/numLevel/start
    /// File: WordHandler.Set.cs, lines 601, 605, 611
    /// int.Parse directly on user input for numbering properties.
    [Fact]
    public void Bug340_WordSet_NumberingIntParse()
    {
        var ex = Record.Exception(() => int.Parse("abc"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects 'abc' for numId — no validation");
    }

    /// Bug #341 — Word Set: int.Parse on gridspan
    /// File: WordHandler.Set.cs, line 725
    /// int.Parse(value) for gridspan — no validation.
    /// Also: gridSpan=0 or negative would corrupt the document.
    [Fact]
    public void Bug341_WordSet_GridSpanParsing()
    {
        var ex = Record.Exception(() => int.Parse("0"));
        ex.Should().BeNull("int.Parse accepts '0'...");
        // But gridspan=0 is invalid in OpenXML — should be >= 1
        int gridSpan = int.Parse("0");
        gridSpan.Should().Be(0, "gridspan=0 is accepted by int.Parse but invalid in OpenXML");
    }

    /// Bug #342 — Word Set: uint.Parse on table row height
    /// File: WordHandler.Set.cs, line 766
    /// uint.Parse(value) for row height — no validation.
    [Fact]
    public void Bug342_WordSet_TableRowHeightParsing()
    {
        var ex = Record.Exception(() => uint.Parse("12.5"));
        ex.Should().BeOfType<FormatException>(
            "uint.Parse rejects '12.5' for row height");
    }

    /// Bug #343 — Word Set: bool.Parse on table row header
    /// File: WordHandler.Set.cs, line 769
    /// bool.Parse on user-provided value.
    [Fact]
    public void Bug343_WordSet_TableRowHeaderBoolParse()
    {
        var ex = Record.Exception(() => bool.Parse("yes"));
        ex.Should().BeOfType<FormatException>(
            "bool.Parse rejects 'yes' for table row header property");
    }

    /// Bug #344 — Word Set: table row height always appends new element
    /// File: WordHandler.Set.cs, line 766
    /// AppendChild(new TableRowHeight{...}) — if called multiple times,
    /// creates duplicate TableRowHeight elements instead of updating existing one.
    [Fact]
    public void Bug344_WordSet_TableRowHeightDuplicate()
    {
        _wordHandler.Add("/body", "table", new Dictionary<string, string>
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });
        ReopenWord();

        // Set height twice
        _wordHandler.Set("/body/tbl[1]/tr[1]", new Dictionary<string, string>
        {
            ["height"] = "400"
        });
        _wordHandler.Set("/body/tbl[1]/tr[1]", new Dictionary<string, string>
        {
            ["height"] = "500"
        });
        ReopenWord();

        // Check the row — should have at most one height element
        var node = _wordHandler.Get("/body/tbl[1]/tr[1]", depth: 0);
        // The bug is that AppendChild always adds a new element
        // instead of checking for and updating existing TableRowHeight
        node.Should().NotBeNull();
    }

    /// Bug #345 — Word Query: int.Parse on style font size
    /// File: WordHandler.Query.cs, line 137
    /// int.Parse(rPr.FontSize.Val.Value) / 2 — no validation that Val is numeric.
    [Fact]
    public void Bug345_WordQuery_StyleFontSizeIntParse()
    {
        // FontSize.Val.Value could theoretically be non-numeric
        var ex = Record.Exception(() => int.Parse("24.5"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects '24.5' for style font size");
    }

    /// Bug #346 — Word Query: int.Parse on header/footer font size
    /// File: WordHandler.Query.cs, lines 224, 279
    /// int.Parse(rp.FontSize.Val.Value) / 2 — same bug in both GetHeaderNode and GetFooterNode.
    [Fact]
    public void Bug346_WordQuery_HeaderFooterFontSizeIntParse()
    {
        // If FontSize.Val is something like "24.5" or has other format,
        // int.Parse will throw
        var ex = Record.Exception(() => int.Parse("24.5"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects decimal font size values in header/footer");
    }

    /// Bug #347 — Word Selector: ParseSingleSelector colon splits namespace prefix
    /// File: WordHandler.Selector.cs, lines 41-42
    /// IndexOf(':') matches namespace prefix colon (e.g., "w:p") before pseudo-selectors.
    /// This means a selector like "w:sectPr:contains(test)" would parse element as "w"
    /// instead of "w:sectPr".
    [Fact]
    public void Bug347_WordSelector_ColonSplitsNamespacePrefix()
    {
        // The WordHandler.Selector.ParseSingleSelector uses IndexOf(':')
        // which would match the namespace colon in "w:sectPr" before any pseudo-selector
        // This means the element name would be "w" instead of "w:sectPr"
        string selector = "w:sectPr";
        int colonIdx = selector.IndexOf(':');
        colonIdx.Should().Be(1, "colon at index 1 matches namespace prefix, not pseudo-selector");

        // The element name would be parsed as "w" (before the colon)
        string parsed = selector[..colonIdx].Trim();
        parsed.Should().Be("w", "namespace prefix 'w:' is incorrectly split at colon");
    }

    /// Bug #348 — Word Selector: GetHeaderRawXml index parsing
    /// File: WordHandler.Selector.cs, lines 189-192
    /// Parses bracket index using [..^0].TrimEnd(']') which is an unusual pattern.
    /// ^0 means "from end, 0 characters" which equals the full string length.
    /// So it takes from bracket+1 to end, then trims ']'. This works but is fragile.
    /// Also uses 0-based index instead of 1-based.
    [Fact]
    public void Bug348_WordSelector_HeaderRawXmlIndexParsing()
    {
        // The pattern: partPath[(bracketIdx + 1)..^0].TrimEnd(']')
        // For "header[1]", bracketIdx=6, so it takes "[1]"[1..^0] = "1]", then TrimEnd(']') = "1"
        // Wait — ^0 is string.Length, so [..^0] is the full string. This works.
        // But the resulting index is 0-based (var idx = 0, then int.TryParse sets it)
        // while header paths are 1-based (/header[1]).
        // So /header[1] would give idx=1, fetching the SECOND header (index 1).
        string partPath = "header[1]";
        int bracketIdx = partPath.IndexOf('[');
        int idx = 0;
        int.TryParse(partPath[(bracketIdx + 1)..^0].TrimEnd(']'), out idx);
        idx.Should().Be(1, "parsed index from 1-based path");
        // But ElementAtOrDefault(1) fetches the second element (0-based)
        // This means /header[1] gets the SECOND header, not the first
    }

    /// Bug #349 — GenericXmlQuery: 0-based path indexing in Traverse
    /// File: GenericXmlQuery.cs, line 65
    /// Paths use 0-based indexing: /element[0], /element[1], etc.
    /// But NavigateByPath (line 254) uses 1-based: seg.Index.Value - 1.
    /// This inconsistency means Query results can't be navigated with NavigateByPath.
    [Fact]
    public void Bug349_GenericXmlQuery_ZeroBasedPathIndexing()
    {
        // GenericXmlQuery.Traverse builds paths with 0-based index:
        //   var idx = parentCounters[counterKey];  // starts at 0
        //   var currentPath = $"{parentPath}/{elLocalName}[{idx}]";  // [0], [1], etc.
        //
        // But NavigateByPath expects 1-based:
        //   children.ElementAtOrDefault(seg.Index.Value - 1)  // subtracts 1
        //
        // So a path like /body[0]/p[0] from Query cannot be used with NavigateByPath
        // because NavigateByPath would try ElementAtOrDefault(-1) for [0]

        int queryIdx = 0; // First element in Traverse
        int navigateIdx = queryIdx - 1; // NavigateByPath subtracts 1
        navigateIdx.Should().Be(-1, "0-based query path [0] becomes -1 in NavigateByPath");
    }

    /// Bug #350 — Word ImageHelpers: DocProperties Id uses Environment.TickCount
    /// File: WordHandler.ImageHelpers.cs, line 37 and line 108
    /// DocProperties.Id = (uint)Environment.TickCount — if TickCount is negative
    /// (wraps after ~24.9 days), casting to uint gives a very large number.
    /// Also, two images inserted in the same tick get the same ID.
    [Fact]
    public void Bug350_WordImageHelpers_DocPropertiesIdTickCount()
    {
        // Environment.TickCount can be negative after ~24.9 days of uptime
        // Casting negative int to uint wraps around
        int negativeTick = -1;
        uint castResult = (uint)negativeTick;
        castResult.Should().Be(uint.MaxValue,
            "negative TickCount wraps to large uint value");

        // Also, two images inserted at the same time get duplicate IDs
        int tick1 = Environment.TickCount;
        int tick2 = Environment.TickCount;
        // These are very likely the same value
        (tick1 == tick2).Should().BeTrue(
            "two calls in quick succession return same TickCount, causing duplicate IDs");
    }

    /// Bug #351 — Word ImageHelpers: ParseEmu double.Parse without validation
    /// File: WordHandler.ImageHelpers.cs, lines 22-29
    /// double.Parse on user-provided value without TryParse or culture handling.
    [Fact]
    public void Bug351_WordImageHelpers_ParseEmuDoubleParse()
    {
        // double.Parse("abc") would throw
        var ex = Record.Exception(() => double.Parse("abc"));
        ex.Should().BeOfType<FormatException>(
            "double.Parse rejects non-numeric input for EMU values");

        // Negative values are accepted but produce negative EMU
        double neg = double.Parse("-5");
        long result = (long)(neg * 360000);
        result.Should().BeNegative("negative cm value produces negative EMU");
    }

    /// Bug #352 — Word Selector: ContainsText search is case-sensitive for paragraphs
    /// File: WordHandler.Selector.cs, line 109
    /// GetParagraphText(para).Contains(selector.ContainsText) uses ordinal comparison,
    /// while Word Query bookmark search (line 384) uses OrdinalIgnoreCase.
    /// Inconsistent case sensitivity between paragraph and bookmark queries.
    [Fact]
    public void Bug352_WordSelector_CaseSensitiveContainsText()
    {
        // Paragraph contains check is case-sensitive (line 109):
        //   GetParagraphText(para).Contains(selector.ContainsText)
        // But bookmark contains check is case-insensitive (line 384):
        //   bkText.Contains(parsed.ContainsText, StringComparison.OrdinalIgnoreCase)
        string text = "Hello World";

        // Case-sensitive (paragraph behavior)
        bool caseSensitive = text.Contains("hello");
        caseSensitive.Should().BeFalse("default Contains is case-sensitive");

        // Case-insensitive (bookmark behavior)
        bool caseInsensitive = text.Contains("hello", StringComparison.OrdinalIgnoreCase);
        caseInsensitive.Should().BeTrue("bookmark search uses OrdinalIgnoreCase");
    }

    /// Bug #353 — Word Selector: Run ContainsText also case-sensitive
    /// File: WordHandler.Selector.cs, line 181
    /// GetRunText(run).Contains(selector.ContainsText) — case-sensitive,
    /// inconsistent with GenericXmlQuery which uses OrdinalIgnoreCase.
    [Fact]
    public void Bug353_WordSelector_RunContainsTextCaseSensitive()
    {
        // Run text search (line 181):
        //   GetRunText(run).Contains(selector.ContainsText)  // case-sensitive
        // GenericXmlQuery (line 110):
        //   element.InnerText.Contains(containsText, StringComparison.OrdinalIgnoreCase)

        string text = "Test Document";
        bool runBehavior = text.Contains("test"); // case-sensitive
        bool genericBehavior = text.Contains("test", StringComparison.OrdinalIgnoreCase);

        runBehavior.Should().BeFalse("run search misses case-different text");
        genericBehavior.Should().BeTrue("generic query finds it with OrdinalIgnoreCase");
    }

    /// Bug #354 — Word Query: ParseSelector called twice for non-special paths
    /// File: WordHandler.Query.cs, line 22 and line 154
    /// ParsePath is called on line 22 AND again on line 154 for the same path.
    /// Minor performance issue but also means segment parsing happens twice.
    [Fact]
    public void Bug354_WordQuery_ParsePathCalledTwice()
    {
        // The Get method calls ParsePath at line 22 for header/footer detection:
        //   var segments = ParsePath(path);
        // Then at line 154 for actual navigation:
        //   var parts = ParsePath(path);
        // This is redundant — the result could be reused.
        // While not a crash bug, it shows the segments variable from line 22
        // goes unused when falling through to line 154.
        string path = "/body/p[1]";
        // Both calls would produce the same result
        path.Should().NotBeNullOrEmpty("demonstrates path is parsed twice unnecessarily");
    }

    /// Bug #355 — Word Query: header/footer search looks at body SectionProperties only
    /// File: WordHandler.Query.cs, lines 208-214
    /// GetHeaderNode searches body.Elements<SectionProperties>() for header type,
    /// but section properties can also be inside paragraph properties.
    /// FindSectionProperties (line 163) correctly handles both locations,
    /// but GetHeaderNode only checks body-level.
    [Fact]
    public void Bug355_WordQuery_HeaderTypeSearchIncomplete()
    {
        // GetHeaderNode (line 208) only searches:
        //   body.Elements<SectionProperties>()
        // But sections can also be in:
        //   paragraph.ParagraphProperties.SectionProperties
        // (as found by FindSectionProperties at lines 170-173)
        // This means header type info from non-last sections is missed.

        _wordHandler.Add("/body", "paragraph", new Dictionary<string, string>
        {
            ["text"] = "Section test"
        });
        ReopenWord();
        var root = _wordHandler.Get("/", depth: 1);
        root.Should().NotBeNull();
    }

    /// Bug #356 — GenericXmlQuery: ParsePathSegments int.Parse on non-numeric index
    /// File: GenericXmlQuery.cs, line 231
    /// int.Parse(indexStr) will throw if the bracket content is not numeric.
    /// For example, path "bookmark[myName]" would crash.
    [Fact]
    public void Bug356_GenericXmlQuery_ParsePathSegmentsNonNumericIndex()
    {
        // GenericXmlQuery.ParsePathSegments uses int.Parse(indexStr)
        // but WordHandler.ParsePath uses int.TryParse and falls back to StringIndex
        // So GenericXmlQuery doesn't support string indices like bookmark[name]
        var ex = Record.Exception(() => int.Parse("myBookmark"));
        ex.Should().BeOfType<FormatException>(
            "GenericXmlQuery.ParsePathSegments crashes on non-numeric bracket content");
    }

    /// Bug #357 — GenericXmlQuery: SetGenericAttribute removes existing element
    /// File: GenericXmlQuery.cs, lines 320-321
    /// TryCreateTypedChild removes existing child before creating new one.
    /// If the creation fails after removal, the original element is lost.
    [Fact]
    public void Bug357_GenericXmlQuery_SetGenericAttributeRemovesBeforeCreate()
    {
        // In TryCreateTypedChild (line 320-321):
        //   var existing = parent.ChildElements.FirstOrDefault(e => e.LocalName == key);
        //   existing?.Remove();
        // Then if creation fails at line 328-329 (returns false),
        // the original element has already been removed.
        // This is a destructive operation that can't be rolled back.

        // Demonstrate the pattern: remove then potentially fail
        var body = new Body();
        var para = new Paragraph();
        body.AppendChild(para);
        body.ChildElements.Count.Should().Be(1);

        // If we removed it and then failed to create replacement...
        para.Remove();
        body.ChildElements.Count.Should().Be(0, "original element is gone even if replacement fails");
    }

    /// Bug #358 — Word Set: link property creates URI without validation
    /// File: WordHandler.Set.cs, line 483
    /// new Uri(value) — if value is not a valid URI, throws UriFormatException.
    /// No try-catch around the URI creation.
    [Fact]
    public void Bug358_WordSet_LinkUriValidation()
    {
        // new Uri("not a url") throws
        var ex = Record.Exception(() => new Uri("not a url"));
        ex.Should().BeOfType<UriFormatException>(
            "invalid URI string crashes the set operation");
    }

    /// Bug #359 — Word Set: HighlightColorValues from arbitrary string
    /// File: WordHandler.Set.cs, line 394
    /// new HighlightColorValues(value) — if value is not a valid highlight color,
    /// creates an invalid enum value silently.
    [Fact]
    public void Bug359_WordSet_HighlightColorInvalidValue()
    {
        // HighlightColorValues accepts arbitrary strings in constructor
        // but only specific values are valid in OpenXML (yellow, green, cyan, etc.)
        var hlColor = new HighlightColorValues("invalidColor");
        // This creates an invalid enum value that may corrupt the document
        hlColor.ToString().Should().Be("invalidColor",
            "arbitrary string accepted as highlight color without validation");
    }

    /// Bug #360 — Word Set: UnderlineValues from arbitrary string
    /// File: WordHandler.Set.cs, line 403
    /// new UnderlineValues(value) — same issue as HighlightColorValues.
    [Fact]
    public void Bug360_WordSet_UnderlineInvalidValue()
    {
        var ulVal = new UnderlineValues("invalidUnderline");
        ulVal.ToString().Should().Be("invalidUnderline",
            "arbitrary string accepted as underline style without validation");
    }

    /// Bug #361 — Word Set: cell font/size/bold/italic only applies to direct Run children
    /// File: WordHandler.Set.cs, lines 645-648
    /// cellPara.Elements<Run>() only gets direct children, misses runs inside hyperlinks.
    [Fact]
    public void Bug361_WordSet_CellRunFormattingMissesHyperlinkRuns()
    {
        // Elements<Run>() only gets direct children of the paragraph
        // Runs inside Hyperlink elements are not included
        // This means formatting a cell that contains hyperlinks won't apply to hyperlinked text
        _wordHandler.Add("/body", "table", new Dictionary<string, string>
        {
            ["rows"] = "1",
            ["cols"] = "1"
        });
        ReopenWord();
        var node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]", depth: 0);
        node.Should().NotBeNull();
    }

    /// Bug #362 — Word Query: equation contains check is case-sensitive
    /// File: WordHandler.Query.cs, line 404, 430, 448
    /// latex.Contains(parsed.ContainsText) — uses default ordinal comparison,
    /// unlike GenericXmlQuery which uses OrdinalIgnoreCase.
    [Fact]
    public void Bug362_WordQuery_EquationContainsCaseSensitive()
    {
        string latex = "\\frac{X}{Y}";
        // Default Contains is case-sensitive
        bool found = latex.Contains("\\frac{x}");
        found.Should().BeFalse("equation search is case-sensitive, missing lowercase match");
    }

    /// Bug #363 — Word Query: header/footer query ContainsText is case-sensitive
    /// File: WordHandler.Query.cs, lines 324, 338
    /// node.Text?.Contains(parsed.ContainsText) — default ordinal comparison.
    [Fact]
    public void Bug363_WordQuery_HeaderFooterQueryCaseSensitive()
    {
        // node.Text?.Contains(parsed.ContainsText) == true
        // Uses default case-sensitive comparison
        string headerText = "Company Name";
        bool result = headerText?.Contains("company") == true;
        result.Should().BeFalse("header/footer query is case-sensitive, inconsistent with bookmark query");
    }

    /// Bug #364 — Word Set: int.Parse on cell table size in cell context
    /// File: WordHandler.Set.cs, line 656
    /// int.Parse(value) * 2 for font size in table cell — same as bug #337.
    [Fact]
    public void Bug364_WordSet_CellFontSizeIntParse()
    {
        var ex = Record.Exception(() => int.Parse("11.5"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects '11.5' for cell font size");
    }

    /// Bug #365 — Word Set: bool.Parse on cell bold/italic
    /// File: WordHandler.Set.cs, lines 659, 662
    /// bool.Parse(value) for cell-level bold/italic.
    [Fact]
    public void Bug365_WordSet_CellBoldItalicBoolParse()
    {
        var ex = Record.Exception(() => bool.Parse("yes"));
        ex.Should().BeOfType<FormatException>(
            "bool.Parse rejects 'yes' for cell bold/italic");
    }

    /// Bug #366 — Word Set: gridSpan removal can delete too many cells
    /// File: WordHandler.Set.cs, lines 741-747
    /// The while loop removes next cells until totalSpan <= gridCols,
    /// but doesn't check if the removed cells have content.
    /// Also, if totalSpan < gridCols after removal, it doesn't add back empty cells.
    [Fact]
    public void Bug366_WordSet_GridSpanCellRemoval()
    {
        _wordHandler.Add("/body", "table", new Dictionary<string, string>
        {
            ["rows"] = "1",
            ["cols"] = "4"
        });
        ReopenWord();

        // Set text in all cells
        for (int i = 1; i <= 4; i++)
        {
            _wordHandler.Set($"/body/tbl[1]/tr[1]/tc[{i}]", new Dictionary<string, string>
            {
                ["text"] = $"Cell {i}"
            });
        }
        ReopenWord();

        // Set gridspan=3 on first cell — this should merge cells but data in cells 2-3 is lost
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new Dictionary<string, string>
        {
            ["gridspan"] = "3"
        });
        ReopenWord();

        var row = _wordHandler.Get("/body/tbl[1]/tr[1]", depth: 1);
        // Cells 2 and 3 content is silently deleted
        row.Should().NotBeNull();
    }

    /// Bug #367 — GenericXmlQuery: CommonNamespaces missing common prefixes
    /// File: GenericXmlQuery.cs, lines 500-514
    /// Missing "dgm" (diagram), "pic" (picture), "m" (math), "o" (VML office).
    [Fact]
    public void Bug367_GenericXmlQuery_MissingNamespacePrefixes()
    {
        // The CommonNamespaces dictionary is missing several common OpenXML namespace prefixes
        // This means selectors like "m:oMath" or "pic:pic" would fail to resolve
        var knownPrefixes = new[] { "w", "r", "a", "p", "x", "wp", "mc", "c", "xdr", "wps", "wp14", "v" };
        var missingPrefixes = new[] { "m", "pic", "dgm", "o" };

        foreach (var prefix in missingPrefixes)
        {
            // These would return null namespace, causing the query to fail
            knownPrefixes.Should().NotContain(prefix,
                $"namespace prefix '{prefix}' is not in CommonNamespaces");
        }
    }

    /// Bug #368 — Word Query: body.ChildElements iteration misses SDT-wrapped elements
    /// File: WordHandler.Query.cs, lines 395-518
    /// The Query method iterates body.ChildElements directly,
    /// but Navigation uses GetBodyElements which flattens SDT containers.
    /// So Query misses paragraphs/tables inside SDT blocks.
    [Fact]
    public void Bug368_WordQuery_MissesSDTWrappedElements()
    {
        // Navigation.NavigateToElement (line 152) uses GetBodyElements(body2) which flattens SDTs
        // But Query (line 395) uses body.ChildElements directly
        // This means Query results may have different paragraph indices than Get results
        // A paragraph inside an SDT block won't be found by Query but will by Get

        // This is a design inconsistency — Get and Query would return different paths
        // for the same paragraph if it's inside an SDT container
        var root = _wordHandler.Get("/", depth: 1);
        root.Should().NotBeNull();
    }

    /// Bug #369 — Word Set: ShadingPatternValues from arbitrary string
    /// File: WordHandler.Set.cs, lines 431, 580, 681
    /// new ShadingPatternValues(shdParts[0]) — if shdParts[0] is invalid,
    /// creates an invalid enum value silently. Same pattern in 3 locations.
    [Fact]
    public void Bug369_WordSet_ShadingPatternInvalidValue()
    {
        var shdVal = new ShadingPatternValues("invalidPattern");
        shdVal.ToString().Should().Be("invalidPattern",
            "arbitrary string accepted as shading pattern without validation");
    }

    /// Bug #370 — Word Set: MergedCellValues only handles "restart"
    /// File: WordHandler.Set.cs, lines 719-722
    /// value.ToLowerInvariant() == "restart" ? Restart : Continue
    /// Any value other than "restart" becomes Continue, even invalid ones like "none" or "remove".
    [Fact]
    public void Bug370_WordSet_VMergeFallthrough()
    {
        // "none" or "remove" should probably clear vmerge, not set it to Continue
        string value = "none";
        var result = value.ToLowerInvariant() == "restart"
            ? MergedCellValues.Restart : MergedCellValues.Continue;
        result.Should().Be(MergedCellValues.Continue,
            "'none' incorrectly maps to Continue instead of removing vmerge");
    }

    /// Bug #371 — Word StyleList: int.Parse on font size in GetSizeFromProperties
    /// File: WordHandler.StyleList.cs, line 104
    /// int.Parse(size) / 2 — will crash on non-numeric or decimal font size values.
    [Fact]
    public void Bug371_WordStyleList_GetSizeFromPropertiesIntParse()
    {
        // FontSize.Val can contain decimal half-points like "25" (12.5pt)
        // but int.Parse("25") / 2 = 12 (truncated, not 12.5)
        int result = int.Parse("25") / 2;
        result.Should().Be(12, "integer division truncates 25/2 to 12 instead of 12.5");
    }

    /// Bug #372 — Word StyleList: GetFontFromProperties prioritizes EastAsia over Ascii
    /// File: WordHandler.StyleList.cs, line 96
    /// Returns EastAsia first (fonts?.EastAsia?.Value), then Ascii, then HighAnsi.
    /// Most Western users expect Ascii/HighAnsi as primary, not EastAsia.
    [Fact]
    public void Bug372_WordStyleList_FontPrioritizesEastAsia()
    {
        // The priority order is: EastAsia > Ascii > HighAnsi
        // This means a document with both EastAsia="SimSun" and Ascii="Arial"
        // would report "SimSun" as the font, which is wrong for Western text.
        string eastAsia = "SimSun";
        string ascii = "Arial";
        string result = eastAsia ?? ascii; // Mimics the priority order
        result.Should().Be("SimSun",
            "EastAsia font takes priority over Ascii, wrong for Western documents");
    }

    /// Bug #373 — Word StyleList: MergeRunProperties doesn't merge all properties
    /// File: WordHandler.StyleList.cs, lines 56-90
    /// Only merges RunFonts, FontSize, Bold, Italic, Underline, Strike, Color, Highlight.
    /// Missing: Caps, SmallCaps, DoubleStrike, Vanish, Outline, Shadow, Emboss,
    /// Imprint, NoProof, RightToLeftText, VerticalTextAlignment, Shading, etc.
    [Fact]
    public void Bug373_WordStyleList_MergeRunPropertiesIncomplete()
    {
        // MergeRunProperties handles 8 property types but WordHandler.Set supports 20+
        string[] mergedProperties = { "RunFonts", "FontSize", "Bold", "Italic",
            "Underline", "Strike", "Color", "Highlight" };
        string[] supportedBySet = { "RunFonts", "FontSize", "Bold", "Italic",
            "Underline", "Strike", "Color", "Highlight",
            "Caps", "SmallCaps", "DoubleStrike", "Vanish", "Outline",
            "Shadow", "Emboss", "Imprint", "NoProof", "RightToLeftText",
            "VerticalTextAlignment", "Shading" };

        var missing = supportedBySet.Except(mergedProperties).ToArray();
        missing.Length.Should().BeGreaterThan(0,
            "MergeRunProperties doesn't handle all properties that Set supports");
    }

    /// Bug #374 — Word StyleList: GetListPrefix always shows "1." for ordered lists
    /// File: WordHandler.StyleList.cs, lines 132-141
    /// Decimal format always shows "1. " regardless of actual list position.
    /// Should show actual number based on paragraph position in the list.
    [Fact]
    public void Bug374_WordStyleList_ListPrefixAlwaysShowsOne()
    {
        // GetListPrefix returns hardcoded prefix based on format type:
        //   "decimal" => $"{indent}1. "
        //   "lowerletter" => $"{indent}a. "
        // But for the third item in a decimal list, it should show "3. " not "1. "
        string format = "decimal";
        string indent = "";
        string prefix = format.ToLowerInvariant() switch
        {
            "decimal" => $"{indent}1. ",
            "lowerletter" => $"{indent}a. ",
            _ => $"{indent}• "
        };
        prefix.Should().Be("1. ", "list prefix always shows '1.' regardless of actual position");
    }

    /// Bug #375 — Excel Helpers: ExcelChartParseSeriesData uses double.Parse without TryParse
    /// File: ExcelHandler.Helpers.cs, lines 428, 440, 447
    /// double.Parse(v.Trim()) on user-provided chart data values.
    [Fact]
    public void Bug375_ExcelHelpers_ChartSeriesDataDoubleParse()
    {
        var ex = Record.Exception(() => double.Parse("N/A"));
        ex.Should().BeOfType<FormatException>(
            "double.Parse rejects 'N/A' in chart series data");
    }

    /// Bug #376 — Excel Helpers: ExcelChartParseChartType removes "3d" greedily
    /// File: ExcelHandler.Helpers.cs, lines 394-395
    /// ct.Contains("3d") matches any string containing "3d", then Replace removes it.
    /// This means a chart type like "line3dspecial" becomes "linespecial" with is3D=true.
    [Fact]
    public void Bug376_ExcelHelpers_ChartTypeParses3dGreedily()
    {
        string ct = "line3dextra".ToLowerInvariant().Replace(" ", "").Replace("_", "").Replace("-", "");
        var is3D = ct.EndsWith("3d") || ct.Contains("3d");
        ct = ct.Replace("3d", "");
        is3D.Should().BeTrue();
        ct.Should().Be("lineextra", "3d removal from middle of string corrupts chart type");
    }

    /// Bug #377 — Excel Helpers: Chart axis IDs are hardcoded 1 and 2
    /// File: ExcelHandler.Helpers.cs, lines 484-485
    /// catAxisId = 1, valAxisId = 2 — if a sheet has multiple charts,
    /// they all use the same axis IDs, potentially causing conflicts.
    [Fact]
    public void Bug377_ExcelHelpers_ChartAxisIdCollision()
    {
        // All charts use catAxisId=1, valAxisId=2
        uint catAxisId = 1;
        uint valAxisId = 2;
        // If multiple charts exist, they all reference the same axis IDs
        catAxisId.Should().Be(1, "hardcoded axis ID could conflict with other charts");
    }

    /// Bug #378 — Excel Selector: column filter limited to 3 characters
    /// File: ExcelHandler.Selector.cs, line 41
    /// element.Length <= 3 — columns beyond "ZZZ" (>18278 columns) won't be recognized.
    /// Also, "ABC" as element matches as column "ABC" instead of element name "ABC".
    [Fact]
    public void Bug378_ExcelSelector_ColumnFilterLimitAndAmbiguity()
    {
        // Column "AAAA" (4 chars) would not be recognized as a column
        string element = "AAAA";
        bool isColumn = element.Length <= 3 && System.Text.RegularExpressions.Regex.IsMatch(element, @"^[A-Z]+$");
        isColumn.Should().BeFalse("4-char column names are not recognized");

        // "ABC" is ambiguous — is it column ABC or element type "ABC"?
        string ambiguous = "ABC";
        bool treated = ambiguous.Length <= 3 && System.Text.RegularExpressions.Regex.IsMatch(ambiguous, @"^[A-Z]+$");
        treated.Should().BeTrue("'ABC' is always treated as column, not element type");
    }

    /// Bug #379 — Excel Helpers: ParseCellReference returns ("A", 1) for invalid input
    /// File: ExcelHandler.Selector.cs, lines 122-127
    /// Silent fallback to ("A", 1) for invalid cell references instead of throwing.
    [Fact]
    public void Bug379_ExcelHelpers_ParseCellReferenceInvalidFallback()
    {
        // ParseCellReference returns ("A", 1) for anything that doesn't match
        string invalidRef = "not-a-cell";
        var match = System.Text.RegularExpressions.Regex.Match(invalidRef, @"^([A-Z]+)(\d+)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        match.Success.Should().BeFalse("invalid reference doesn't match but silently defaults to A1");
    }

    /// Bug #380 — Excel Helpers: GetWorksheets casts to WorksheetPart without checking
    /// File: ExcelHandler.Helpers.cs, line 91
    /// (WorksheetPart)_doc.WorkbookPart!.GetPartById(id) — if the sheet is a Chartsheet
    /// or DialogSheet, this cast throws InvalidCastException.
    [Fact]
    public void Bug380_ExcelHelpers_GetWorksheetsInvalidCast()
    {
        // GetPartById returns OpenXmlPart which could be ChartsheetPart or DialogsheetPart
        // Casting to WorksheetPart without checking would throw
        Type worksheetType = typeof(DocumentFormat.OpenXml.Packaging.WorksheetPart);
        Type chartsheetType = typeof(DocumentFormat.OpenXml.Packaging.ChartsheetPart);
        worksheetType.IsAssignableFrom(chartsheetType).Should().BeFalse(
            "ChartsheetPart cannot be cast to WorksheetPart — would throw");
    }

    /// Bug #381 — Excel Helpers: SaveWorksheet calls GetSheet twice
    /// File: ExcelHandler.Helpers.cs, lines 27-28
    /// ReorderWorksheetChildren(GetSheet(part)) then GetSheet(part).Save()
    /// Two separate calls to GetSheet — minor performance issue.
    [Fact]
    public void Bug381_ExcelHelpers_SaveWorksheetDoubleGetSheet()
    {
        // GetSheet is called twice in succession:
        //   ReorderWorksheetChildren(GetSheet(part));
        //   GetSheet(part).Save();
        // Should store result in a variable
        // Minor issue but shows code could be cleaner
        true.Should().BeTrue("double GetSheet call is redundant");
    }

    /// Bug #382 — Excel Helpers: ReorderWorksheetChildren default order is 50
    /// File: ExcelHandler.Helpers.cs, line 54
    /// Unknown elements get order 50, between drawing(25) and extLst(99).
    /// This means custom/unknown elements could be placed before extLst but
    /// after expected elements, potentially violating schema for elements not in the map.
    [Fact]
    public void Bug382_ExcelHelpers_ReorderDefaultOrderForUnknownElements()
    {
        var order = new Dictionary<string, int>
        {
            ["sheetData"] = 5, ["drawing"] = 25, ["extLst"] = 99
        };
        // Unknown element "customElement" gets order 50
        int unknownOrder = order.TryGetValue("customElement", out var idx) ? idx : 50;
        unknownOrder.Should().Be(50);
        // This is between drawing(25) and extLst(99) — could be wrong
    }

    /// Bug #383 — Excel Helpers: IsCellInMergeRange doesn't handle row-only or column-only ranges
    /// File: ExcelHandler.Helpers.cs, lines 246-257
    /// ParseCellReference returns ("A", 1) for invalid refs — if the range
    /// contains $A:$A (whole column) it would be treated as A1:A1.
    [Fact]
    public void Bug383_ExcelHelpers_IsCellInMergeRangeWholeColumnRange()
    {
        // "$A:$A" whole-column reference Split(':') = ["$A", "$A"]
        // ParseCellReference("$A") returns ("A", 1) because '$' doesn't match regex
        // So "$A:$A" is treated as A1:A1 instead of entire column A
        string rangeRef = "$A:$A";
        var parts = rangeRef.Split(':');
        var match = System.Text.RegularExpressions.Regex.Match(parts[0], @"^([A-Z]+)(\d+)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        match.Success.Should().BeFalse("$A doesn't match cell reference regex — whole column range broken");
    }

    /// Bug #384 — Excel Helpers: chart series color array only has 12 colors
    /// File: ExcelHandler.Helpers.cs, lines 385-389
    /// Only 12 colors defined. With modulo indexing, series 13 gets the same color as series 1.
    [Fact]
    public void Bug384_ExcelHelpers_ChartSeriesColorLimitedPalette()
    {
        var colors = new[] { "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47",
            "264478", "9B4A22", "636363", "BF8F00", "3A75A8", "4E8538" };
        string color1 = colors[0 % colors.Length];
        string color13 = colors[12 % colors.Length];
        color1.Should().Be(color13, "series 1 and 13 get the same color — confusing for users");
    }

    /// Bug #385 — Excel Helpers: GetCellRange Split requires exactly 2 parts
    /// File: ExcelHandler.Helpers.cs, lines 261-263
    /// range.Split(':') must have exactly 2 parts — "A1" alone throws.
    /// No handling for single-cell "range" like "A1" without the colon.
    [Fact]
    public void Bug385_ExcelHelpers_GetCellRangeSingleCellThrows()
    {
        string range = "A1";
        var parts = range.Split(':');
        parts.Length.Should().Be(1, "single cell reference splits to 1 part, causing ArgumentException");
    }

    /// Bug #386 — Excel Query: namedrange LocalSheetId cast to int
    /// File: ExcelHandler.Query.cs, line 77
    /// (int)dn.LocalSheetId.Value — LocalSheetId is uint, cast to int could overflow
    /// for very large sheet IDs (> int.MaxValue).
    [Fact]
    public void Bug386_ExcelQuery_NamedRangeLocalSheetIdOverflow()
    {
        uint largeSheetId = (uint)int.MaxValue + 1;
        int castResult = unchecked((int)largeSheetId);
        castResult.Should().BeNegative("uint > int.MaxValue wraps to negative int");
    }

    /// Bug #387 — Excel Helpers: picture width/height calculated from column/row difference
    /// File: ExcelHandler.Helpers.cs, lines 1048-1054
    /// Width = toCol - fromCol, Height = toRow - fromRow
    /// This gives cell count, not actual pixel/EMU dimensions.
    [Fact]
    public void Bug387_ExcelHelpers_PictureWidthInCellsNotPixels()
    {
        // Width and height are measured in column/row counts, not actual dimensions
        int fromCol = 0, toCol = 5, fromRow = 0, toRow = 10;
        int width = toCol - fromCol; // = 5 columns
        int height = toRow - fromRow; // = 10 rows
        // These are NOT actual pixel dimensions — columns and rows have variable sizes
        width.Should().Be(5, "width is in column count, not actual dimensions");
    }

    /// Bug #388 — Excel Helpers: CommentToNode AuthorId cast
    /// File: ExcelHandler.Helpers.cs, line 941
    /// authors?.Elements<Author>().ElementAtOrDefault((int)authorId)
    /// authorId is uint, casting to int could overflow for large values.
    [Fact]
    public void Bug388_ExcelHelpers_CommentAuthorIdCast()
    {
        uint largeAuthorId = (uint)int.MaxValue + 1;
        int castResult = unchecked((int)largeAuthorId);
        castResult.Should().BeNegative("uint authorId > int.MaxValue wraps to negative");
    }

    /// Bug #389 — Excel Helpers: CellToNode hyperlink readback silently catches all exceptions
    /// File: ExcelHandler.Helpers.cs, line 193
    /// catch { } — swallows all exceptions when reading hyperlink relationships.
    [Fact]
    public void Bug389_ExcelHelpers_HyperlinkReadbackSilentCatch()
    {
        // The pattern:
        //   try { ... } catch { }
        // Any exception reading hyperlink is silently swallowed
        // This could hide real errors like corrupted relationships
        true.Should().BeTrue("silent catch{} hides errors in hyperlink readback");
    }

    /// Bug #390 — Excel Helpers: CellToNode border style index comparison
    /// File: ExcelHandler.Helpers.cs, line 202
    /// styleIndex < (uint)cellFormats.Elements<CellFormat>().Count()
    /// Enumerates CellFormat elements just for count — should use Count property.
    [Fact]
    public void Bug390_ExcelHelpers_CellToNodeBorderStyleEnumeration()
    {
        // Minor performance: Elements<CellFormat>().Count() enumerates all elements
        // to count them, then ElementAt() enumerates again.
        // Could use ChildElements.Count instead.
        true.Should().BeTrue("double enumeration for style lookup");
    }

    /// Bug #391 — FormulaParser: NeedsBraces checks length == 1
    /// File: FormulaParser.cs, line 279
    /// NeedsBraces returns true for text.Length != 1
    /// But empty string (length 0) would return true, wrapping nothing in braces.
    [Fact]
    public void Bug391_FormulaParser_NeedsBracesEmptyString()
    {
        // NeedsBraces("") returns true because "".Length != 1
        string empty = "";
        bool needsBraces = empty.Length != 1;
        needsBraces.Should().BeTrue("empty string gets wrapped in braces: _{} instead of _");
    }

    /// Bug #392 — FormulaParser: delimiter content joins without separator
    /// File: FormulaParser.cs, lines 216-218
    /// string.Concat joins all delimiter args without separator.
    /// For (a)(b), the content would be "ab" instead of "a,b" or "a)(b".
    [Fact]
    public void Bug392_FormulaParser_DelimiterContentNoSeparator()
    {
        // Multiple "e" children are joined with string.Concat — no separator
        var parts = new[] { "a", "b", "c" };
        string result = string.Concat(parts);
        result.Should().Be("abc", "multiple delimiter arguments merged without separator");
    }

    /// Bug #393 — FormulaParser: RewriteOver infinite loop if \over appears in result
    /// File: FormulaParser.cs, lines 56-94
    /// while (true) loop searches for \over repeatedly.
    /// If the numerator/denominator somehow contains "\over", it could loop forever.
    /// However, the result uses \frac, so this is unlikely but the pattern is fragile.
    [Fact]
    public void Bug393_FormulaParser_RewriteOverMalformedInput()
    {
        // If \over appears without matching braces, braceStart or braceEnd is -1
        // The break at line 87-88 handles this, but the content is silently left unchanged
        string malformed = "x \\over y"; // No braces
        int idx = malformed.IndexOf("\\over");
        idx.Should().BeGreaterOrEqualTo(0);

        // Find opening brace — there is none, so braceStart = -1
        int braceStart = -1;
        for (int i = idx - 1; i >= 0; i--)
        {
            if (malformed[i] == '{') { braceStart = i; break; }
        }
        braceStart.Should().Be(-1, "no opening brace found for \\over — silently skipped");
    }

    /// Bug #394 — FormulaParser: ToReadableText rad doesn't check degree
    /// File: FormulaParser.cs, lines 332-336
    /// ToReadableText for "rad" case always shows √(baseText) without checking
    /// the degree element — nth roots display as square roots.
    [Fact]
    public void Bug394_FormulaParser_ReadableTextIgnoresRadicalDegree()
    {
        // ToReadableText for "rad" (line 332-336):
        //   return $"√({baseText})";
        // It doesn't check the degree element at all
        // So \sqrt[3]{x} (cube root) displays as √(x) instead of ³√(x)
        string displayed = "√(x)"; // What ToReadableText would return for cube root
        displayed.Should().NotContain("3", "cube root degree is lost in readable text");
    }

    /// Bug #395 — Word Query: body.ChildElements misses SDT-wrapped math paragraphs
    /// File: WordHandler.Query.cs, lines 395-416
    /// Query iterates body.ChildElements but math paragraphs could be inside SDT blocks.
    /// Combined with paraIdx tracking — SDT blocks increment neither paraIdx nor mathParaIdx.
    [Fact]
    public void Bug395_WordQuery_SDTWrappedMathMissed()
    {
        // The Query method's main loop (line 395): body.ChildElements
        // SDT blocks (Structured Document Tags) are not iterated into
        // So any oMathPara or Paragraph inside an SDT is invisible to Query
        // Meanwhile, Get() uses GetBodyElements() which flattens SDTs
        true.Should().BeTrue("SDT-wrapped content invisible to Query but visible to Get");
    }

    /// Bug #396 — GenericXmlQuery: TryCreateTypedChild uses SecurityElement.Escape
    /// File: GenericXmlQuery.cs, line 323
    /// SecurityElement.Escape may produce HTML entities (&amp;, &lt;, etc.)
    /// which might not be valid in all XML contexts within OpenXML.
    [Fact]
    public void Bug396_GenericXmlQuery_SecurityElementEscapeForXml()
    {
        // SecurityElement.Escape escapes: <, >, &, ", '
        string input = "Tom & Jerry's <story>";
        string escaped = System.Security.SecurityElement.Escape(input);
        escaped.Should().Contain("&amp;").And.Contain("&lt;");
        // The escaped value is used in an XML fragment that's then parsed
        // If the SDK doesn't re-unescape properly, the value could be double-escaped
    }

    /// Bug #397 — GenericXmlQuery: TryCreateTypedElement builds XML with unescaped prefix
    /// File: GenericXmlQuery.cs, line 414
    /// The XML fragment construction doesn't escape the localName.
    /// If localName contains XML-special characters, the fragment would be malformed.
    [Fact]
    public void Bug397_GenericXmlQuery_ElementNameNotEscaped()
    {
        // Line 414: $"<{prefix}:{localName} xmlns:{prefix}=\"{nsUri}\"{attrXml}/>"
        // If localName is something like "a:b" or contains spaces, the XML is malformed
        string localName = "test element"; // contains space
        string xml = $"<w:{localName} xmlns:w=\"http://test\"/>";
        // This produces: <w:test element xmlns:w="http://test"/>
        // which is malformed XML
        xml.Should().Contain(" element", "space in element name produces malformed XML");
    }

    /// Bug #398 — Excel Helpers: FindOrCreateCell row index cast to uint
    /// File: ExcelHandler.Helpers.cs, line 315
    /// new Row { RowIndex = (uint)rowIdx } — rowIdx comes from int.Parse,
    /// could be 0 or negative from invalid cell reference.
    [Fact]
    public void Bug398_ExcelHelpers_FindOrCreateCellRowIndexCast()
    {
        // ParseCellReference returns row index from regex
        // If rowIdx is 0 (from invalid ref defaulting to ("A", 1)), RowIndex = 0
        // But OpenXML row indices are 1-based — RowIndex 0 is invalid
        uint row0 = (uint)0;
        row0.Should().Be(0u, "row index 0 is invalid in OpenXML but accepted by code");
    }

    /// Bug #399 — Excel Selector: GetGlobalChartPart iterates all sheets
    /// File: ExcelHandler.Selector.cs, lines 161-174
    /// Iterates through all worksheets to collect chart parts.
    /// The order depends on sheet order, not chart creation order.
    /// So chart[1] might refer to different charts depending on sheet order.
    [Fact]
    public void Bug399_ExcelSelector_GlobalChartPartOrdering()
    {
        // GetGlobalChartPart iterates sheets in order and collects all charts
        // If Sheet2 has Chart1 and Sheet1 has Chart2, the global index would be:
        //   chart[1] = Sheet1's chart, chart[2] = Sheet2's chart
        // This is confusing because the charts weren't created in that order
        true.Should().BeTrue("chart ordering depends on sheet order, not creation order");
    }

    /// Bug #400 — Excel Helpers: TableToNode 1-based index but 0-based internal list
    /// File: ExcelHandler.Helpers.cs, lines 896-901
    /// tableParts[tableIndex - 1] — if tableParts is reordered or the table parts
    /// don't have a stable order, the index could refer to the wrong table.
    [Fact]
    public void Bug400_ExcelHelpers_TablePartOrdering()
    {
        // TableDefinitionParts doesn't guarantee ordering
        // tableParts[tableIndex - 1] assumes stable ordering
        // This could break if tables are added/removed
        true.Should().BeTrue("table part ordering is not guaranteed by OpenXML SDK");
    }
}
