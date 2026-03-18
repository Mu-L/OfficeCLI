// Bug hunt Part 20 — deeper property gaps, Set/Get asymmetries, validation issues.

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart20 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public BugHuntPart20()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt20_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt20_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt20_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_pptxPath);
        using (var pptx = new PowerPointHandler(_pptxPath, editable: true))
            pptx.Add("/", "slide", null, new());
        _wordHandler = new WordHandler(_docxPath, editable: true);
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
    }

    public void Dispose()
    {
        _wordHandler.Dispose();
        _excelHandler.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private WordHandler ReopenWord()
    {
        _wordHandler.Dispose();
        _wordHandler = new WordHandler(_docxPath, editable: true);
        return _wordHandler;
    }

    private ExcelHandler ReopenExcel()
    {
        _excelHandler.Dispose();
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
        return _excelHandler;
    }


    // ==================== BUG #1: Word table Set style not reflected in Get ====================
    // Set style on table, but table Get doesn't report the style name.
    [Fact]
    public void Word_Table_SetStyle_RoundTrip()
    {
        _wordHandler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        _wordHandler.Set("/body/tbl[1]", new()
        {
            ["style"] = "TableGrid"
        });

        ReopenWord();

        var table = _wordHandler.Get("/body/tbl[1]");
        table.Format.Should().ContainKey("style",
            "table style should persist and be readable after reopen");
    }


    // ==================== BUG #2: PPTX shape color readback format inconsistency ====================
    // Set uses bare hex "FF0000", but Get reads from RgbColorModelHex which stores without #.
    // What does Get actually return?
    [Fact]
    public void Pptx_Shape_Color_SetGet_FormatConsistency()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Red text",
            ["color"] = "#FF0000"
        });

        var shape = pptx.Get("/slide[1]/shape[1]");
        shape.Format.Should().ContainKey("color");

        var color = shape.Format["color"]?.ToString();
        // Color should be consistent: either always with # or always without
        color.Should().Be("FF0000",
            "color readback should be bare hex without # prefix, matching the Set format");
    }


    // ==================== BUG #3: Excel row height not in Get ====================
    [Fact]
    public void Excel_Row_Height_RoundTrip()
    {
        _excelHandler.Set("/Sheet1/row[1]", new()
        {
            ["height"] = "30"
        });

        var row = _excelHandler.Get("/Sheet1/row[1]");
        row.Should().NotBeNull();

        row.Format.Should().ContainKey("height",
            "row Get should include height when it's been set");
    }


    // ==================== BUG #4: Word paragraph widowcontrol not in Get ====================
    [Fact]
    public void Word_Paragraph_Get_ShouldIncludeWidowControl()
    {
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Widow test",
            ["widowcontrol"] = "true"
        });

        var para = _wordHandler.Get("/body/p[1]");
        para.Format.Should().ContainKey("widowcontrol",
            "paragraph Get should expose widowcontrol when it's set");
    }


    // ==================== BUG #5: PPTX table Get missing row count when queried at depth=0 ====================
    [Fact]
    public void Pptx_Table_Get_Depth0_ShouldIncludeRowCount()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "3",
            ["cols"] = "2"
        });

        var table = pptx.Get("/slide[1]/table[1]", depth: 0);
        table.Should().NotBeNull();
        table.Format.Should().ContainKey("rows");
        table.Format["rows"].Should().Be(3,
            "table should report correct row count even at depth=0");
    }


    // ==================== BUG #6: Word Set paragraph alignment then Get shows enum name ====================
    // Set uses "center", but Get might return "Center" (capitalized enum name).
    // The values should be consistent for round-trip comparison.
    [Fact]
    public void Word_Paragraph_Alignment_ValueFormat()
    {
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Centered",
            ["alignment"] = "center"
        });

        var para = _wordHandler.Get("/body/p[1]");
        para.Format.Should().ContainKey("alignment");

        var alignment = para.Format["alignment"]?.ToString();
        // Set uses "center" (lowercase), Get should return the same format
        alignment.Should().Be("center",
            "alignment value should be lowercase 'center' to match the Set input format, " +
            "not the .NET enum name 'Center'");
    }


    // ==================== BUG #7: Excel cell type readback not matching Set ====================
    [Fact]
    public void Excel_Cell_Type_RoundTrip()
    {
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Hello",
            ["type"] = "string"
        });

        var cell = _excelHandler.Get("/Sheet1/A1");
        cell.Should().NotBeNull();

        cell.Format.Should().ContainKey("type",
            "cell Get should include the data type when it's been explicitly set");
    }


    // ==================== BUG #8: PPTX shape fill readback after Set ====================
    [Fact]
    public void Pptx_Shape_Fill_RoundTrip()
    {
        using var pptx = new PowerPointHandler(_pptxPath, editable: true);

        pptx.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Filled",
            ["fill"] = "00FF00"
        });

        var shape = pptx.Get("/slide[1]/shape[1]");
        shape.Format.Should().ContainKey("fill");

        var fill = shape.Format["fill"]?.ToString();
        fill.Should().Be("00FF00",
            "shape fill should round-trip: Set '00FF00' → Get '00FF00'");
    }


    // ==================== BUG #9: Word section Get doesn't include orientation ====================
    [Fact]
    public void Word_Section_Get_ShouldIncludeOrientation()
    {
        _wordHandler.Set("/section[1]", new()
        {
            ["orientation"] = "landscape"
        });

        var section = _wordHandler.Get("/section[1]");
        section.Should().NotBeNull();

        section.Format.Should().ContainKey("orientation",
            "section Get should include orientation when it's been set");
    }


    // ==================== BUG #10: Excel hidden column not in Get ====================
    [Fact]
    public void Excel_Column_Hidden_RoundTrip()
    {
        _excelHandler.Set("/Sheet1/col[2]", new()
        {
            ["hidden"] = "true"
        });

        var col = _excelHandler.Get("/Sheet1/col[2]");
        col.Should().NotBeNull();

        col.Format.Should().ContainKey("hidden",
            "column Get should include hidden property when it's been set");
    }


    // ==================== BUG #11: Word paragraph shading not in Get ====================
    [Fact]
    public void Word_Paragraph_Shading_RoundTrip()
    {
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Shaded",
            ["shading"] = "FFFF00"
        });

        var para = _wordHandler.Get("/body/p[1]");
        para.Format.Should().ContainKey("shading",
            "paragraph Get should include shading when it's been set");
    }
}
