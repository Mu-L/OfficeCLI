// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for PPTX: each test creates a blank file, adds elements,
/// queries them, and modifies them — exercising the full Create→Add→Get→Set lifecycle.
/// </summary>
public class PptxFunctionalTests : IDisposable
{
    private readonly string _path;
    private PowerPointHandler _handler;

    public PptxFunctionalTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_path);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    // Reopen the file to verify persistence
    private PowerPointHandler Reopen()
    {
        _handler.Dispose();
        _handler = new PowerPointHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Slide lifecycle ====================

    [Fact]
    public void AddSlide_ReturnsPath_Slide1()
    {
        var path = _handler.Add("/", "slide", null, new Dictionary<string, string>());
        path.Should().Be("/slide[1]");
    }

    [Fact]
    public void AddSlide_Get_ReturnsSlideType()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var node = _handler.Get("/slide[1]");
        node.Type.Should().Be("slide");
    }

    [Fact]
    public void AddSlide_Multiple_PathIncrements()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var path3 = _handler.Add("/", "slide", null, new Dictionary<string, string>());
        path3.Should().Be("/slide[3]");
    }

    [Fact]
    public void AddSlide_WithTitle_TitleVisibleInText()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "Hello World" });
        var node = _handler.Get("/slide[1]", depth: 2);
        var allText = node.Children.SelectMany(c => c.Children).Select(c => c.Text).Concat(
                      node.Children.Select(c => c.Text))
                      .Where(t => t != null).ToList();
        allText.Any(t => t!.Contains("Hello World")).Should().BeTrue();
    }

    // ==================== Shape lifecycle ====================

    [Fact]
    public void AddShape_ReturnsPath_Shape1()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var shapePath = _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Test" });
        shapePath.Should().Be("/slide[1]/shape[1]");
    }

    [Fact]
    public void AddShape_WithText_TextIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Hello Shape" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Hello Shape");
    }

    [Fact]
    public void AddShape_WithFill_FillIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Filled", ["fill"] = "FF0000" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("fill");
        node.Format["fill"].Should().Be("FF0000");
    }

    [Fact]
    public void AddShape_WithPosition_PositionIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Positioned",
            ["x"] = "2cm",
            ["y"] = "3cm",
            ["width"] = "5cm",
            ["height"] = "2cm"
        });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format["x"].Should().Be("2cm");
        node.Format["y"].Should().Be("3cm");
        node.Format["width"].Should().Be("5cm");
        node.Format["height"].Should().Be("2cm");
    }

    [Fact]
    public void AddShape_WithName_NameIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Named", ["name"] = "MyBox" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format["name"].Should().Be("MyBox");
    }

    // ==================== Set: modify shape properties ====================

    [Fact]
    public void SetShape_Bold_BoldIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Normal" });

        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["bold"] = "true" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("bold");
        node.Format["bold"].Should().Be(true);
    }

    [Fact]
    public void SetShape_Italic_ItalicIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Normal" });

        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["italic"] = "true" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("italic");
        node.Format["italic"].Should().Be(true);
    }

    [Fact]
    public void SetShape_Fill_FillIsUpdated()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "A", ["fill"] = "0000FF" });

        // Change fill from blue to red
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["fill"] = "FF0000" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format["fill"].Should().Be("FF0000");
    }

    [Fact]
    public void SetShape_FontSize_SizeIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Big Text" });

        // size property accepts a raw point number (stored as pt*100 internally)
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["size"] = "24" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("size");
        node.Format["size"].Should().Be("24pt");
    }

    [Fact]
    public void SetShape_Position_PositionIsUpdated()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string> { ["text"] = "A" });

        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string>
        {
            ["x"] = "4cm",
            ["y"] = "5cm"
        });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format["x"].Should().Be("4cm");
        node.Format["y"].Should().Be("5cm");
    }

    // ==================== Query ====================

    [Fact]
    public void QueryShapes_AfterAddTwo_ReturnsBoth()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string> { ["text"] = "A" });
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string> { ["text"] = "B" });

        var nodes = _handler.Query("shape");
        nodes.Should().HaveCountGreaterThanOrEqualTo(2);
    }

    [Fact]
    public void GetRoot_AfterAddThreeSlides_HasThreeChildren()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/", "slide", null, new Dictionary<string, string>());

        var root = _handler.Get("/");
        root.Children.Should().HaveCount(3);
        root.Children.Should().AllSatisfy(c => c.Type.Should().Be("slide"));
    }

    // ==================== Table lifecycle ====================

    [Fact]
    public void AddTable_ReturnsTablePath()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var path = _handler.Add("/slide[1]", "table", null,
            new Dictionary<string, string> { ["rows"] = "2", ["cols"] = "3" });
        path.Should().Be("/slide[1]/table[1]");
    }

    [Fact]
    public void AddTable_Get_HasCorrectDimensions()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "table", null,
            new Dictionary<string, string> { ["rows"] = "3", ["cols"] = "4" });

        var node = _handler.Get("/slide[1]/table[1]");
        node.Type.Should().Be("table");
        node.Format["rows"].Should().Be(3);
        node.Format["cols"].Should().Be(4);
    }

    [Fact]
    public void SetTableCell_TextIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "table", null,
            new Dictionary<string, string> { ["rows"] = "2", ["cols"] = "2" });

        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]",
            new Dictionary<string, string> { ["text"] = "Cell A1" });

        var table = _handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = table.Children
            .FirstOrDefault(r => r.Type == "tr")
            ?.Children.FirstOrDefault(c => c.Type == "tc");
        cell.Should().NotBeNull();
        cell!.Text.Should().Be("Cell A1");
    }

    // ==================== Slide background ====================

    [Fact]
    public void AddSlide_WithBackground_BackgroundIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "FF0000" });

        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        node.Format["background"].Should().Be("FF0000");
    }

    [Fact]
    public void AddSlide_WithGradientBackground_BackgroundIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "FF0000-0000FF" });

        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        var bg = node.Format["background"]?.ToString();
        bg.Should().NotBeNull();
        bg!.Should().Contain("FF0000");
        bg.Should().Contain("0000FF");
    }

    [Fact]
    public void SetSlideBackground_SolidColor_IsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Set("/slide[1]", new Dictionary<string, string> { ["background"] = "FF0000" });

        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        node.Format["background"].Should().Be("FF0000");
    }

    [Fact]
    public void SetSlideBackground_UpdateColor_NewColorIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "FF0000" });

        _handler.Set("/slide[1]", new Dictionary<string, string> { ["background"] = "0000FF" });

        var node = _handler.Get("/slide[1]");
        node.Format["background"].Should().Be("0000FF");
    }

    [Fact]
    public void SetSlideBackground_None_RemovesBackground()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "FF0000" });
        _handler.Set("/slide[1]", new Dictionary<string, string> { ["background"] = "none" });

        var node = _handler.Get("/slide[1]");
        node.Format.Should().NotContainKey("background");
    }

    [Fact]
    public void AddSlide_WithBackground_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "00FF00" });

        var handler2 = Reopen();
        var node = handler2.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        node.Format["background"].Should().Be("00FF00");
    }

    // ==================== Persistence ====================

    [Fact]
    public void AddShape_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Persistent" });

        var handler2 = Reopen();
        var node = handler2.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Persistent");
    }
}
