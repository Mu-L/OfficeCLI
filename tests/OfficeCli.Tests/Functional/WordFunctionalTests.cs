// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for DOCX: each test creates a blank file, adds elements,
/// queries them, and modifies them — exercising the full Create→Add→Get→Set lifecycle.
/// </summary>
public class WordFunctionalTests : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public WordFunctionalTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private WordHandler Reopen()
    {
        _handler.Dispose();
        _handler = new WordHandler(_path, editable: true);
        return _handler;
    }

    // ==================== DOCX Hyperlinks ====================

    [Fact]
    public void Hyperlink_Lifecycle()
    {
        // 1. Add paragraph + hyperlink
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>());
        var path = _handler.Add("/body/p[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://first.com",
            ["text"] = "Click here"
        });
        path.Should().Be("/body/p[1]/hyperlink[1]");

        // 2. Get + Verify type, url, text
        var node = _handler.Get("/body/p[1]/hyperlink[1]");
        node.Type.Should().Be("hyperlink");
        node.Text.Should().Be("Click here");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://first.com");

        // 3. Verify paragraph text contains link text
        var para = _handler.Get("/body/p[1]");
        para.Text.Should().Contain("Click here");

        // 4. Query + Verify
        var results = _handler.Query("hyperlink");
        results.Should().Contain(n => n.Type == "hyperlink" && n.Text == "Click here");

        // 5. Set (update URL via run) + Verify
        _handler.Set("/body/p[1]/r[1]", new Dictionary<string, string> { ["link"] = "https://updated.com" });
        node = _handler.Get("/body/p[1]/hyperlink[1]");
        ((string)node.Format["link"]).Should().StartWith("https://updated.com");
    }

    // ==================== DOCX Numbering / Lists ====================

    [Fact]
    public void ListStyle_Bullet_Lifecycle()
    {
        // 1. Add paragraph with bullet list
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Bullet item 1",
            ["liststyle"] = "bullet"
        });

        // 2. Get + Verify all numbering properties
        var node = _handler.Get("/body/p[1]");
        node.Text.Should().Be("Bullet item 1");
        node.Format.Should().ContainKey("numid");
        node.Format.Should().ContainKey("numlevel");
        node.Format.Should().ContainKey("listStyle");
        node.Format.Should().ContainKey("numFmt");
        node.Format.Should().ContainKey("start");
        ((int)node.Format["numlevel"]).Should().Be(0);
        ((string)node.Format["listStyle"]).Should().Be("bullet");
        ((string)node.Format["numFmt"]).Should().Be("bullet");
        ((int)node.Format["start"]).Should().Be(1);

        // 3. Set — change numlevel
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["numlevel"] = "1" });

        // 4. Get + Verify level changed
        node = _handler.Get("/body/p[1]");
        ((int)node.Format["numlevel"]).Should().Be(1);

        // 5. Persist + Verify
        var handler2 = Reopen();
        node = handler2.Get("/body/p[1]");
        node.Text.Should().Be("Bullet item 1");
        ((string)node.Format["listStyle"]).Should().Be("bullet");
        ((int)node.Format["numlevel"]).Should().Be(1);
    }

    [Fact]
    public void ListStyle_Ordered_Lifecycle()
    {
        // 1. Add paragraph with ordered list
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Step 1",
            ["liststyle"] = "numbered"
        });

        // 2. Get + Verify
        var node = _handler.Get("/body/p[1]");
        node.Text.Should().Be("Step 1");
        node.Format.Should().ContainKey("numid");
        node.Format.Should().ContainKey("listStyle");
        node.Format.Should().ContainKey("numFmt");
        ((string)node.Format["listStyle"]).Should().Be("ordered");
        ((string)node.Format["numFmt"]).Should().Be("decimal");

        // 3. Set — change to bullet
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["liststyle"] = "bullet" });

        // 4. Get + Verify changed
        node = _handler.Get("/body/p[1]");
        ((string)node.Format["listStyle"]).Should().Be("bullet");

        // 5. Persist + Verify
        var handler2 = Reopen();
        node = handler2.Get("/body/p[1]");
        ((string)node.Format["listStyle"]).Should().Be("bullet");
    }

    [Fact]
    public void ListStyle_None_RemovesNumbering()
    {
        // 1. Add paragraph with bullet list
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Will lose numbering",
            ["liststyle"] = "bullet"
        });
        var node = _handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("numid");

        // 2. Set listStyle=none to remove numbering
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["liststyle"] = "none" });

        // 3. Get + Verify numbering removed
        node = _handler.Get("/body/p[1]");
        node.Text.Should().Be("Will lose numbering");
        node.Format.Should().NotContainKey("numid");
        node.Format.Should().NotContainKey("listStyle");

        // 4. Persist + Verify still removed
        var handler2 = Reopen();
        node = handler2.Get("/body/p[1]");
        node.Format.Should().NotContainKey("numid");
    }

    [Fact]
    public void ListStyle_Continuation_SharesNumId()
    {
        // 1. Add first bullet paragraph — creates new numbering
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Item A",
            ["liststyle"] = "bullet"
        });
        var numId1 = (int)_handler.Get("/body/p[1]").Format["numid"];

        // 2. Add second consecutive bullet paragraph — should reuse same numId
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Item B",
            ["liststyle"] = "bullet"
        });
        var numId2 = (int)_handler.Get("/body/p[2]").Format["numid"];

        numId2.Should().Be(numId1, "consecutive same-type list items should share numId");

        // 3. Persist + Verify continuation survives reopen
        var handler2 = Reopen();
        var n1 = handler2.Get("/body/p[1]");
        var n2 = handler2.Get("/body/p[2]");
        ((int)n1.Format["numid"]).Should().Be((int)n2.Format["numid"]);
    }

    [Fact]
    public void ListStyle_StartValue_Lifecycle()
    {
        // 1. Add ordered list starting from 5
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Step 5",
            ["liststyle"] = "numbered",
            ["start"] = "5"
        });

        // 2. Get + Verify start value
        var node = _handler.Get("/body/p[1]");
        node.Format.Should().ContainKey("start");
        ((int)node.Format["start"]).Should().Be(5);
        ((string)node.Format["listStyle"]).Should().Be("ordered");

        // 3. Set — change start value via Set
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["start"] = "10" });

        // 4. Get + Verify
        node = _handler.Get("/body/p[1]");
        ((int)node.Format["start"]).Should().Be(10);

        // 5. Persist + Verify
        var handler2 = Reopen();
        node = handler2.Get("/body/p[1]");
        ((int)node.Format["start"]).Should().Be(10);
    }

    [Fact]
    public void ListStyle_NumId_RawAccess()
    {
        // 1. Add paragraph with listStyle
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Raw item",
            ["liststyle"] = "bullet"
        });

        // 2. Get the numid back
        var numId = (int)_handler.Get("/body/p[1]").Format["numid"];
        numId.Should().BeGreaterThan(0);

        // 3. Add another paragraph using the raw numid
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Same list",
            ["numid"] = numId.ToString(),
            ["numlevel"] = "0"
        });

        // 4. Get + Verify shared numid
        var node2 = _handler.Get("/body/p[2]");
        ((int)node2.Format["numid"]).Should().Be(numId);
        ((int)node2.Format["numlevel"]).Should().Be(0);

        // 5. Persist + Verify
        var handler2 = Reopen();
        node2 = handler2.Get("/body/p[2]");
        ((int)node2.Format["numid"]).Should().Be(numId);
    }

    [Fact]
    public void ListStyle_NineLevels_Supported()
    {
        // 1. Add a bullet list
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Deep nesting",
            ["liststyle"] = "bullet"
        });

        // 2. Set numlevel to 8 (0-based, 9th level)
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["numlevel"] = "8" });

        // 3. Get + Verify level 8 works
        var node = _handler.Get("/body/p[1]");
        ((int)node.Format["numlevel"]).Should().Be(8);
        ((string)node.Format["listStyle"]).Should().Be("bullet");

        // 4. Persist + Verify
        var handler2 = Reopen();
        node = handler2.Get("/body/p[1]");
        ((int)node.Format["numlevel"]).Should().Be(8);
    }

    [Fact]
    public void ListStyle_NumFmt_ReturnsSpecificFormat()
    {
        // 1. Add ordered list
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Level 0",
            ["liststyle"] = "numbered"
        });

        // 2. Verify level 0 = decimal
        var node = _handler.Get("/body/p[1]");
        ((string)node.Format["numFmt"]).Should().Be("decimal");

        // 3. Set to level 1
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["numlevel"] = "1" });

        // 4. Verify level 1 = lowerLetter
        node = _handler.Get("/body/p[1]");
        ((string)node.Format["numFmt"]).Should().Be("lowerLetter");

        // 5. Set to level 2
        _handler.Set("/body/p[1]", new Dictionary<string, string> { ["numlevel"] = "2" });

        // 6. Verify level 2 = lowerRoman
        node = _handler.Get("/body/p[1]");
        ((string)node.Format["numFmt"]).Should().Be("lowerRoman");
    }

    [Fact]
    public void ListStyle_Query_FilterByListStyle()
    {
        // 1. Add mixed paragraphs
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Normal paragraph"
        });
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Bullet item",
            ["liststyle"] = "bullet"
        });
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Ordered item",
            ["liststyle"] = "numbered"
        });

        // 2. Query + Verify filtering
        var bullets = _handler.Query("paragraph[liststyle=bullet]");
        bullets.Should().HaveCount(1);
        bullets[0].Text.Should().Be("Bullet item");

        var ordered = _handler.Query("paragraph[liststyle=ordered]");
        ordered.Should().HaveCount(1);
        ordered[0].Text.Should().Be("Ordered item");

        // 3. Query by numid
        var numId = (int)_handler.Get("/body/p[2]").Format["numid"];
        var byNumId = _handler.Query($"paragraph[numid={numId}]");
        byNumId.Should().ContainSingle(n => n.Text == "Bullet item");
    }

    [Fact]
    public void Hyperlink_Persist_SurvivesReopenFile()
    {
        _handler.Add("/body", "paragraph", null, new Dictionary<string, string>());
        _handler.Add("/body/p[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://original.com",
            ["text"] = "My link"
        });
        _handler.Set("/body/p[1]/r[1]", new Dictionary<string, string> { ["link"] = "https://persist.com" });

        var handler2 = Reopen();
        var node = handler2.Get("/body/p[1]/hyperlink[1]");
        node.Text.Should().Be("My link");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://persist.com");
    }
}
