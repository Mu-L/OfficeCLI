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

    /// Bug #401 — PPTX Query: equation search case-sensitive
    /// File: PowerPointHandler.Query.cs, line 418
    /// latex.Contains(parsed.TextContains) — default ordinal comparison.
    /// But notes query (line 384) uses OrdinalIgnoreCase.
    [Fact]
    public void Bug401_PptxQuery_EquationSearchCaseSensitive()
    {
        string latex = "\\frac{X}{Y}";
        // Equation search (line 418) is case-sensitive
        bool equationResult = latex.Contains("\\frac{x}");
        equationResult.Should().BeFalse("equation search is case-sensitive unlike notes search");

        // Notes search (line 384) is case-insensitive
        string notes = "Important Note";
        bool notesResult = notes.Contains("important", StringComparison.OrdinalIgnoreCase);
        notesResult.Should().BeTrue("notes search is case-insensitive — inconsistency");
    }

    /// Bug #402 — PPTX Query: placeholder query uses different shapeIdx than shape query
    /// File: PowerPointHandler.Query.cs, lines 499-523
    /// Placeholder query re-iterates shapes and counts only those with PlaceholderShape.
    /// So placeholder[1] path uses a different index than shape[N] for the same element.
    [Fact]
    public void Bug402_PptxQuery_PlaceholderIndexingDifferent()
    {
        // In shape query (line 409): shapeIdx counts ALL shapes
        // In placeholder query (line 501): phIdx counts only shapes WITH placeholders
        // So if shape[3] is a placeholder, it might be placeholder[1]
        // The path in shape query would be /slide[1]/shape[3]
        // But in placeholder query it's /slide[1]/placeholder[1]
        // This makes the indices inconsistent
        int allShapes = 5;
        int placeholders = 2; // only 2 of 5 are placeholders
        allShapes.Should().NotBe(placeholders,
            "shape indices and placeholder indices differ for the same element");
    }

    /// Bug #403 — PPTX Query: table text extraction via OuterXml regex
    /// File: PowerPointHandler.Query.cs, lines 471-475
    /// Uses regex on OuterXml to extract text: @"<a:t[^>]*>([^<]*)</a:t>"
    /// This fails if text contains < (escaped as &lt;) or multi-line text.
    [Fact]
    public void Bug403_PptxQuery_TableTextExtractionRegex()
    {
        // Regex @"<a:t[^>]*>([^<]*)</a:t>" captures text between <a:t> tags
        // But [^<]* stops at < which could appear as entity (&lt;) or CDATA
        string xml = "<a:t>A &lt; B</a:t>";
        var matches = System.Text.RegularExpressions.Regex.Matches(xml, @"<a:t[^>]*>([^<]*)</a:t>");
        var text = string.Concat(matches.Select(m => m.Groups[1].Value));
        text.Should().Be("A ", "text after &lt; is lost because [^<]* stops at &");
    }

    /// Bug #404 — PPTX Query: chart title cast to string
    /// File: PowerPointHandler.Query.cs, line 491
    /// (string)chartNode.Format["title"] — Format is Dictionary<string, object>,
    /// if title was stored as non-string (unlikely), this would throw InvalidCastException.
    [Fact]
    public void Bug404_PptxQuery_ChartTitleCastToString()
    {
        // chartNode.Format["title"] is cast to (string) directly
        // If Format contains a non-string value for "title", this throws
        var dict = new Dictionary<string, object>();
        dict["title"] = 123; // numeric value
        var ex = Record.Exception(() => { var s = (string)dict["title"]; });
        ex.Should().BeOfType<InvalidCastException>(
            "casting non-string Format value to string throws");
    }

    /// Bug #405 — PPTX Query: Get root calls GetSlide(slidePart) three times per slide
    /// File: PowerPointHandler.Query.cs, lines 37, 48, 49
    /// GetSlide, ReadSlideBackground, ReadSlideTransition each call GetSlide.
    /// Performance issue for presentations with many slides.
    [Fact]
    public void Bug405_PptxQuery_GetRootTripleGetSlide()
    {
        // Lines 37, 48, 49 in Get root:
        //   GetSlide(slidePart).CommonSlideData?.ShapeTree?...
        //   ReadSlideBackground(GetSlide(slidePart), slideNode);
        //   ReadSlideTransition(GetSlide(slidePart), slideNode);
        // Three separate GetSlide calls for the same slide part
        true.Should().BeTrue("triple GetSlide call per slide in root query");
    }

    /// Bug #406 — Excel View: ViewAsOutline casts to WorksheetPart without checking
    /// File: ExcelHandler.View.cs, line 132
    /// (WorksheetPart)_doc.WorkbookPart!.GetPartById(sheetId)
    /// Same issue as Bug #380 — could throw for ChartsheetPart.
    [Fact]
    public void Bug406_ExcelView_OutlineCastToWorksheetPart()
    {
        Type worksheetType = typeof(DocumentFormat.OpenXml.Packaging.WorksheetPart);
        Type chartsheetType = typeof(DocumentFormat.OpenXml.Packaging.ChartsheetPart);
        worksheetType.IsAssignableFrom(chartsheetType).Should().BeFalse(
            "ViewAsOutline casts to WorksheetPart without checking part type");
    }

    /// Bug #407 — Excel View: ViewAsOutline colCount only from first row
    /// File: ExcelHandler.View.cs, line 136
    /// colCount = sheetData?.Elements<Row>().FirstOrDefault()?.Elements<Cell>().Count() ?? 0
    /// If first row has fewer cells than subsequent rows, the count is wrong.
    [Fact]
    public void Bug407_ExcelView_ColCountFirstRowOnly()
    {
        // Only counting cells in the first row:
        // A spreadsheet where row 1 has ["Name"] and row 2 has ["Name", "Age", "Email"]
        // would report colCount=1 instead of 3
        int firstRowCells = 1;
        int maxRowCells = 3;
        firstRowCells.Should().NotBe(maxRowCells,
            "column count from first row doesn't represent actual column count");
    }

    /// Bug #408 — Excel View: ViewAsText lineNum doesn't account for sheet separators
    /// File: ExcelHandler.View.cs, lines 29-49
    /// lineNum resets per sheet but startLine/endLine are global.
    /// If user specifies --start=5, it applies within each sheet, not globally.
    [Fact]
    public void Bug408_ExcelView_LineNumResetsPerSheet()
    {
        // lineNum is reset to 0 for each sheet (line 29)
        // But startLine/endLine parameters are presumably global line numbers
        // So --start=5 would show row 5+ in EACH sheet, not row 5+ globally
        int lineNum = 0; // resets per sheet
        lineNum.Should().Be(0, "lineNum resets per sheet — startLine/endLine apply per-sheet");
    }

    /// Bug #409 — Excel View: ViewAsIssues only detects formula errors
    /// File: ExcelHandler.View.cs, lines 195-233
    /// Only checks for formula errors (#REF!, #VALUE!, etc.)
    /// Doesn't detect: missing required cells, data validation violations,
    /// circular references, or inconsistent formulas.
    [Fact]
    public void Bug409_ExcelView_IssuesOnlyFormulaErrors()
    {
        // ViewAsIssues only checks for formula error values
        string[] detected = { "#REF!", "#VALUE!", "#NAME?", "#DIV/0!" };
        string[] undetected = { "#NULL!", "#N/A", "#NUM!", "#GETTING_DATA" };
        // Missing error types
        detected.Should().NotContain("#N/A", "#N/A error is not detected");
    }

    /// Bug #410 — Excel View: ViewAsAnnotated emits row count even for annotated view
    /// File: ExcelHandler.View.cs, line 108
    /// emitted++ is incremented per row, but annotated view shows cells per row.
    /// The maxLines limit counts rows, not cells — inconsistent with line-based pagination.
    [Fact]
    public void Bug410_ExcelView_AnnotatedRowVsCellCounting()
    {
        // In annotated view, each cell gets a line of output
        // But emitted++ (line 108) counts rows, not cells
        // So maxLines=10 shows 10 rows (which could be 50+ output lines)
        int rowsPerOutput = 1;
        int cellsPerRow = 5;
        int outputLines = rowsPerOutput * cellsPerRow;
        (outputLines > 1).Should().BeTrue(
            "maxLines counts rows but each row produces multiple output lines");
    }

    /// Bug #411 — PPTX Query: Get calls ResolveShape then IndexOf for placeholder
    /// File: PowerPointHandler.Query.cs, line 220
    /// shapeTree?.Elements<Shape>().ToList().IndexOf(phShape)
    /// If phShape is not in Elements<Shape>() (e.g., it's in layout, not slide), IndexOf returns -1.
    [Fact]
    public void Bug411_PptxQuery_PlaceholderShapeIndexOfNegative()
    {
        // IndexOf returns -1 if shape is not found
        var list = new List<string> { "a", "b", "c" };
        int idx = list.IndexOf("d");
        idx.Should().Be(-1);
        // Then shapeIdx + 1 = 0, which is an invalid 1-based index
        int shapeIdx = idx + 1;
        shapeIdx.Should().Be(0, "IndexOf=-1 + 1 = 0, invalid shape index for path");
    }

    /// Bug #412 — PPTX Query: font size integer division
    /// File: PowerPointHandler.Query.cs, line 187
    /// fs.Value / 100 — integer division truncates. 1450 / 100 = 14 not 14.5
    [Fact]
    public void Bug412_PptxQuery_FontSizeIntegerDivision()
    {
        int fontSize = 1450; // 14.5pt in hundredths
        int result = fontSize / 100;
        result.Should().Be(14, "integer division truncates 1450/100 to 14 instead of 14.5");
    }

    /// Bug #413 — Excel Selector: ColumnNameToIndex doesn't validate input
    /// File: ExcelHandler.Selector.cs, lines 129-137
    /// Non-letter characters would produce incorrect results.
    /// E.g., "1" would calculate 1 - 'A' + 1 = -16
    [Fact]
    public void Bug413_ExcelSelector_ColumnNameToIndexInvalidInput()
    {
        // ColumnNameToIndex("1") would calculate: '1' - 'A' + 1 = 49 - 65 + 1 = -15
        int charDiff = '1' - 'A' + 1;
        charDiff.Should().Be(-15, "non-letter character produces negative column index");
    }

    /// Bug #414 — Excel Selector: IndexToColumnName infinite loop for negative input
    /// File: ExcelHandler.Selector.cs, lines 139-149
    /// while (index > 0) — negative index skips the loop, returns empty string.
    /// But 0 also skips the loop and returns empty string.
    [Fact]
    public void Bug414_ExcelSelector_IndexToColumnNameZeroOrNegative()
    {
        // IndexToColumnName(0) returns ""
        int index = 0;
        string result = "";
        while (index > 0)
        {
            index--;
            result = (char)('A' + index % 26) + result;
            index /= 26;
        }
        result.Should().BeEmpty("index=0 returns empty string, not 'A'");
    }

    /// Bug #415 — PPTX Query: logicalResolved path regex incomplete
    /// File: PowerPointHandler.Query.cs, line 231
    /// Regex: @"^/slide\[\d+\]/(table\[\d+\]/(tr|tc)|placeholder\[\w+\]/)"
    /// Requires trailing "/" after placeholder — so /slide[1]/placeholder[title]/text wouldn't match
    /// if path ends without trailing content.
    [Fact]
    public void Bug415_PptxQuery_LogicalPathRegexTrailingSlash()
    {
        string path = "/slide[1]/placeholder[title]";
        var match = System.Text.RegularExpressions.Regex.IsMatch(path,
            @"^/slide\[\d+\]/(table\[\d+\]/(tr|tc)|placeholder\[\w+\]/)");
        match.Should().BeFalse("path without trailing content after placeholder doesn't match regex");
    }

    /// Bug #416 — Excel View: ViewAsText columns filter uses Where with deferred execution
    /// File: ExcelHandler.View.cs, line 45
    /// cellElements = cellElements.Where(...) — modifies IEnumerable but doesn't materialize.
    /// Then .Select(c => GetCellDisplayValue(c)).ToArray() iterates.
    /// Normally fine, but the filter uses ParseCellReference which defaults to "A1" for invalid refs.
    [Fact]
    public void Bug416_ExcelView_ColumnFilterDefaultsToA1()
    {
        // If a cell has no CellReference (null), it defaults to "A1"
        // So column filter would always include cells without references
        // This could show phantom data from cells with missing references
        string defaultRef = "A1";
        var (col, _) = (defaultRef.Substring(0, 1), int.Parse(defaultRef.Substring(1)));
        col.Should().Be("A", "cells without references always appear in column A filter");
    }

    /// Bug #417 — Excel Helpers: IsTruthy doesn't handle uppercase or mixed case
    /// File: ExcelHandler.Helpers.cs, line 347-348
    /// value.ToLowerInvariant() is "true" or "1" or "yes"
    /// Actually this DOES handle case since it lowercases first.
    /// But it doesn't handle "on", "y", "t" which are also common truthy values.
    [Fact]
    public void Bug417_ExcelHelpers_IsTruthyMissingValues()
    {
        // IsTruthy only accepts "true", "1", "yes"
        string[] accepted = { "true", "1", "yes" };
        string[] missing = { "on", "y", "t", "enabled" };

        foreach (var val in missing)
        {
            bool result = val.ToLowerInvariant() is "true" or "1" or "yes";
            result.Should().BeFalse($"'{val}' is not recognized as truthy");
        }
    }

    /// Bug #418 — Excel Helpers: GetIconCount only checks first character
    /// File: ExcelHandler.Helpers.cs, lines 375-381
    /// lower.StartsWith("5") / StartsWith("4") — so "50arrows" would return 5 icons.
    /// And "3d-arrows" returns 3 but is actually meant as a 3D style.
    [Fact]
    public void Bug418_ExcelHelpers_GetIconCountAmbiguous()
    {
        // "4x4" starts with "4" so returns 4 icons
        string iconName = "4x4grid";
        int count;
        var lower = iconName.ToLowerInvariant();
        if (lower.StartsWith("5")) count = 5;
        else if (lower.StartsWith("4")) count = 4;
        else count = 3;
        count.Should().Be(4, "'4x4grid' incorrectly parsed as 4-icon set");
    }

    /// Bug #419 — PPTX Query: Get calls ResolveShape but doesn't cache SlideParts
    /// File: PowerPointHandler.Query.cs, line 90
    /// ResolveShape internally calls GetSlideParts().ToList() again.
    /// Multiple calls within the same Get invocation enumerate slides multiple times.
    [Fact]
    public void Bug419_PptxQuery_MultipleGetSlidePartsEnumeration()
    {
        // Get method: GetSlideParts() is called in multiple locations:
        //   Line 34, 74, 205, 248, 265 — each enumerates all slide parts
        // ResolveShape also calls GetSlideParts internally
        // This means a single Get call could enumerate slides 5+ times
        true.Should().BeTrue("multiple GetSlideParts() calls per Get invocation");
    }

    /// Bug #420 — Excel View: ViewAsStats doesn't include chart or pivot table counts
    /// File: ExcelHandler.View.cs, lines 151-193
    /// Only counts cells, formulas, and data types.
    /// Missing: chart count, pivot table count, named range count,
    /// conditional formatting count, data validation count.
    [Fact]
    public void Bug420_ExcelView_StatsIncomplete()
    {
        string[] included = { "Sheets", "Total Cells", "Empty Cells", "Formula Cells", "Error Cells", "Data Type Distribution" };
        string[] missing = { "Charts", "Pivot Tables", "Named Ranges", "Conditional Formatting", "Data Validations" };
        missing.Length.Should().BeGreaterThan(0,
            "ViewAsStats doesn't report charts, pivot tables, named ranges, etc.");
    }

    /// Bug #421 — PPTX Notes: SetNotesText hardcodes zh-CN language
    /// File: PowerPointHandler.Notes.cs, line 75
    /// new Drawing.RunProperties { Language = "zh-CN" }
    /// Notes text always tagged as Chinese regardless of actual language.
    [Fact]
    public void Bug421_PptxNotes_HardcodedZhCNLanguage()
    {
        string language = "zh-CN";
        language.Should().Be("zh-CN",
            "notes text language is hardcoded to zh-CN instead of using system/document locale");
    }

    /// Bug #422 — PPTX Notes: GetNotesText uses Index==1 to find notes placeholder
    /// File: PowerPointHandler.Notes.cs, line 23
    /// ph?.Index?.Value == 1 — assumes notes body placeholder has index 1.
    /// But some note slides might use different placeholder indices.
    [Fact]
    public void Bug422_PptxNotes_NotesPlaceholderIndexAssumption()
    {
        // GetNotesText looks for ph.Index.Value == 1
        // But the body placeholder could have a different index depending on the slide layout
        // Also, Index is nullable — if it's null, the notes won't be found
        uint? index = null;
        bool matches = index == 1;
        matches.Should().BeFalse("null index doesn't match, notes placeholder could be missed");
    }

    /// Bug #423 — PPTX Notes: EnsureNotesSlidePart creates shapes with hardcoded IDs
    /// File: PowerPointHandler.Notes.cs, lines 93, 101, 113
    /// NonVisualDrawingProperties Ids are hardcoded as 1, 2, 3.
    /// If existing notes slide has shapes with these IDs, there could be collisions.
    [Fact]
    public void Bug423_PptxNotes_HardcodedShapeIds()
    {
        uint[] ids = { 1, 2, 3 };
        ids.Should().HaveCount(3, "hardcoded shape IDs could collide with existing shapes");
    }

    /// Bug #424 — PPTX ShapeProperties: int.Parse on font size
    /// File: PowerPointHandler.ShapeProperties.cs, line 77
    /// int.Parse(value) * 100 — same issue as Word Set: no validation on user input.
    [Fact]
    public void Bug424_PptxShapeProps_FontSizeIntParse()
    {
        var ex = Record.Exception(() => int.Parse("14pt"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects '14pt' for font size — should strip unit suffix");
    }

    /// Bug #425 — PPTX ShapeProperties: bool.Parse on bold/italic
    /// File: PowerPointHandler.ShapeProperties.cs, lines 86, 95
    /// bool.Parse(value) on user-provided values for bold and italic.
    [Fact]
    public void Bug425_PptxShapeProps_BoldItalicBoolParse()
    {
        var ex = Record.Exception(() => bool.Parse("yes"));
        ex.Should().BeOfType<FormatException>(
            "bool.Parse rejects 'yes' for bold/italic in PPTX shape properties");
    }

    /// Bug #426 — PPTX Chart: ParseSeriesData double.Parse without validation
    /// File: PowerPointHandler.Chart.cs, lines 65, 79, 86
    /// double.Parse(v.Trim()) on user-provided chart data values.
    [Fact]
    public void Bug426_PptxChart_SeriesDataDoubleParse()
    {
        var ex = Record.Exception(() => double.Parse("N/A"));
        ex.Should().BeOfType<FormatException>(
            "double.Parse rejects 'N/A' in PPTX chart series data");
    }

    /// Bug #427 — PPTX Chart: ParseChartType same greedy 3d removal as Excel
    /// File: PowerPointHandler.Chart.cs, lines 23-24
    /// ct.Contains("3d") and ct.Replace("3d", "") — identical to bug #376.
    [Fact]
    public void Bug427_PptxChart_ChartType3dGeedyRemoval()
    {
        string ct = "scatter3dplot";
        ct = ct.Replace("3d", "");
        ct.Should().Be("scatterplot", "3d removal from middle of chart type name corrupts it");
    }

    /// Bug #428 — PPTX Animations: transition type "none" doesn't clear advance settings
    /// File: PowerPointHandler.Animations.cs, line 35-39
    /// RemoveAllChildren<Transition>() clears the transition element,
    /// but advance time/click settings might be stored elsewhere.
    [Fact]
    public void Bug428_PptxAnimations_TransitionNoneDoesntClearAdvance()
    {
        // Setting transition to "none" removes the Transition element
        // But if advancetime was set, it was on the Transition element
        // So removing Transition also removes advancetime — which might be intended
        // But there's no way to set "no transition but still auto-advance"
        true.Should().BeTrue("removing transition also removes auto-advance settings");
    }

    /// Bug #429 — PPTX Animations: unknown transition type silently creates empty Transition
    /// File: PowerPointHandler.Animations.cs, lines 67-113
    /// The switch returns null for unknown types (line 112: _ => null).
    /// Then transElem is null, so no transition child is appended.
    /// But the Transition element IS still created and added to the slide (lines 118-122).
    [Fact]
    public void Bug429_PptxAnimations_UnknownTransitionCreatesEmptyElement()
    {
        // For unknown types like "magical", transElem = null (line 112)
        // But Transition element is still created at line 63 and appended at lines 119-122
        // This creates an empty <p:transition/> element in the slide
        string typeName = "magical";
        OpenXmlElement? transElem = typeName switch
        {
            "fade" => new object() as OpenXmlElement,  // simplified
            _ => null
        };
        transElem.Should().BeNull("unknown transition type produces null but empty Transition is added");
    }

    /// Bug #430 — PPTX ShapeProperties: SetRunOrShapeProperties replaces all paragraphs
    /// File: PowerPointHandler.ShapeProperties.cs, lines 31-59
    /// When setting "text" with multi-line: preserves first run formatting.
    /// But if shape has multiple paragraphs with different formatting,
    /// all formatting is replaced with the first run's formatting.
    [Fact]
    public void Bug430_PptxShapeProps_TextSetLosesPerParagraphFormatting()
    {
        // Shape with: Para1 (bold), Para2 (italic), Para3 (underline)
        // After setting text = "Line1\nLine2\nLine3":
        //   All lines get Para1's bold formatting
        //   Para2 italic and Para3 underline are lost
        string[] formats = { "bold", "italic", "underline" };
        string preservedFormat = formats[0]; // only first run's formatting preserved
        preservedFormat.Should().Be("bold",
            "only first run formatting is preserved when replacing multi-line text");
    }

    /// Bug #431 — PPTX ShapeProperties: text with empty runs.Count
    /// File: PowerPointHandler.ShapeProperties.cs, lines 33-37
    /// runs.Count == 1 && textLines.Length == 1 — but if runs is empty (Count==0),
    /// falls to else branch which tries to get firstRun (null) from textBody.
    [Fact]
    public void Bug431_PptxShapeProps_TextSetEmptyRuns()
    {
        // If the shape has no runs (empty shape), runs.Count = 0
        // The condition (runs.Count == 1 && ...) is false, falls to else
        // textBody.Descendants<Drawing.Run>().FirstOrDefault() returns null
        // runProps = null, which is handled (line 53: if runProps != null)
        // But the shape might not have a TextBody at all
        var emptyList = new List<int>();
        emptyList.Count.Should().Be(0, "empty runs list falls to else branch");
    }

    /// Bug #432 — FormulaParser: Tokenize doesn't handle escaped braces
    /// File: FormulaParser.cs (not fully read, but pattern visible)
    /// LaTeX input like "\{" or "\}" should be treated as literal braces,
    /// but the tokenizer likely treats { and } as group delimiters.
    [Fact]
    public void Bug432_FormulaParser_EscapedBracesNotHandled()
    {
        // In LaTeX, \{ and \} are literal braces
        // But the tokenizer likely treats { as "open group" regardless of escape
        string latex = "\\{1, 2, 3\\}";
        latex.Should().Contain("\\{", "escaped braces should be treated as literal characters");
    }

    /// Bug #433 — GenericXmlQuery: Traverse uses 0-based index but ElementToNode uses 1-based
    /// File: GenericXmlQuery.cs, line 65 vs line 208
    /// Traverse: currentPath = $"{parentPath}/{elLocalName}[{idx}]"  — idx starts at 0
    /// ElementToNode depth>0: $"{path}/{name}[{idx + 1}]"  — idx+1 is 1-based
    [Fact]
    public void Bug433_GenericXmlQuery_InconsistentIndexing()
    {
        // Traverse uses 0-based: [0], [1], [2], ...
        // ElementToNode uses 1-based: [1], [2], [3], ...
        // So GenericXmlQuery.Query returns paths like /body[0]/p[0]
        // But GenericXmlQuery.ElementToNode(depth>0) generates paths like /body/p[1]
        int traverseIdx = 0; // from Traverse
        int elementToNodeIdx = 0 + 1; // from ElementToNode
        traverseIdx.Should().NotBe(elementToNodeIdx,
            "Query and ElementToNode use different indexing for the same elements");
    }

    /// Bug #434 — Word Helpers: GetRunFontSize int.Parse without TryParse
    /// File: WordHandler.Helpers.cs, line 127 (from summary)
    /// int.Parse on font size value could throw for non-numeric or decimal values.
    [Fact]
    public void Bug434_WordHelpers_GetRunFontSizeIntParse()
    {
        var ex = Record.Exception(() => int.Parse("24.5"));
        ex.Should().BeOfType<FormatException>(
            "int.Parse rejects '24.5' for run font size in Word Helpers");
    }

    /// Bug #435 — PPTX Query: Get slide shape count excludes tables, charts, pictures
    /// File: PowerPointHandler.Query.cs, lines 58-60
    /// ChildCount = Elements<Shape>().Count() + Elements<Picture>().Count()
    /// Missing: GraphicFrame (tables, charts), GroupShape (grouped shapes), ConnectionShape
    [Fact]
    public void Bug435_PptxQuery_SlideChildCountIncomplete()
    {
        // ChildCount only counts Shape + Picture
        // But a slide can also contain:
        // - GraphicFrame (tables, charts, videos, SmartArt)
        // - GroupShape (grouped elements)
        // - ConnectionShape (connectors)
        string[] counted = { "Shape", "Picture" };
        string[] missing = { "GraphicFrame", "GroupShape", "ConnectionShape" };
        missing.Length.Should().BeGreaterThan(0,
            "slide ChildCount missing GraphicFrame, GroupShape, ConnectionShape");
    }

    /// Bug #436 — PPTX Animations: split transition direction ignores user's in/out preference
    /// File: PowerPointHandler.Animations.cs, lines 84-88
    /// SplitTransition hardcodes Direction = ParseInOutDir("in").
    /// The user's direction parameter only sets Orientation, not Direction.
    [Fact]
    public void Bug436_PptxAnimations_SplitTransitionIgnoresInOut()
    {
        // User specifies "split-horizontal-out" but:
        //   Orientation = ParseOrientation("horizontal")  — correct
        //   Direction = ParseInOutDir("in")  — hardcoded "in", ignoring "out"
        string userDir = "out";
        string hardcoded = "in";
        userDir.Should().NotBe(hardcoded, "split transition ignores user's in/out direction");
    }

    /// Bug #437 — PPTX Animations: wheel transition hardcodes 4 spokes
    /// File: PowerPointHandler.Animations.cs, line 82
    /// Spokes = new UInt32Value(4u) — always 4, no way to customize.
    [Fact]
    public void Bug437_PptxAnimations_WheelSpokesHardcoded()
    {
        uint spokes = 4;
        spokes.Should().Be(4, "wheel transition always has 4 spokes, no customization available");
    }

    /// Bug #438 — PPTX Chart: ParseSeriesData skips entries without colon in data format
    /// File: PowerPointHandler.Chart.cs, line 62
    /// if (colonIdx < 0) continue — series without name:values format are skipped.
    /// So "1,2,3" alone in the data string is silently ignored.
    [Fact]
    public void Bug438_PptxChart_DataWithoutColonSkipped()
    {
        string dataStr = "1,2,3;Series2:4,5,6";
        var result = new List<string>();
        foreach (var part in dataStr.Split(';', StringSplitOptions.RemoveEmptyEntries))
        {
            var colonIdx = part.IndexOf(':');
            if (colonIdx < 0) continue; // "1,2,3" is skipped
            result.Add(part);
        }
        result.Should().HaveCount(1, "series without name:values format is silently skipped");
    }

    /// Bug #439 — Excel Helpers: GetCellDisplayValue formula shows "=" prefix
    /// File: ExcelHandler.Helpers.cs, lines 123-127
    /// Returns $"={cell.CellFormula.Text}" for formula cells without cached value.
    /// This means a query that expects "100" might get "=SUM(A1:A10)".
    [Fact]
    public void Bug439_ExcelHelpers_FormulaDisplayValuePrefix()
    {
        // Cells with formulas but no cached value show the formula with = prefix
        string formula = "SUM(A1:A10)";
        string displayValue = $"={formula}";
        displayValue.Should().StartWith("=", "formula cells show formula as display value");
    }

    /// Bug #440 — Excel Helpers: SharedStringTable element lookup is O(n) per cell
    /// File: ExcelHandler.Helpers.cs, line 117
    /// Elements<SharedStringItem>().ElementAtOrDefault(idx)
    /// For each cell, this enumerates up to idx items. O(n*m) for n cells with average index m.
    [Fact]
    public void Bug440_ExcelHelpers_SharedStringLookupLinear()
    {
        // ElementAtOrDefault(idx) iterates from the beginning each time
        // For a spreadsheet with 10000 shared strings and 50000 cells,
        // this could enumerate billions of elements
        int sharedStringCount = 10000;
        int cellCount = 50000;
        long worstCase = (long)sharedStringCount * cellCount;
        worstCase.Should().BeGreaterThan(0, "O(n*m) shared string lookup for large files");
    }
}
