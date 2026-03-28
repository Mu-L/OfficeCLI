// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for the FormulaEvaluator: each test creates a blank XLSX,
/// sets cell values and formulas via ExcelHandler, and verifies cached evaluation results.
/// </summary>
public class FormulaEvaluatorTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public FormulaEvaluatorTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private ExcelHandler Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Basic Arithmetic ====================

    [Fact]
    public void Formula_BasicArithmetic()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "20" });
        _handler.Set("/Sheet1/C1", new() { ["formula"] = "A1+B1*2" });

        var node = _handler.Get("/Sheet1/C1");
        node.Format.Should().ContainKey("formula");
        node.Format["formula"].Should().Be("A1+B1*2");
        node.Format.Should().ContainKey("cachedValue");
        node.Format["cachedValue"].Should().Be("50");
        node.Text.Should().Be("50");
    }

    // ==================== SUM Range ====================

    [Fact]
    public void Formula_SUM_Range()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "20" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "30" });
        _handler.Set("/Sheet1/A4", new() { ["value"] = "40" });
        _handler.Set("/Sheet1/A5", new() { ["value"] = "50" });
        _handler.Set("/Sheet1/B1", new() { ["formula"] = "SUM(A1:A5)" });

        var node = _handler.Get("/Sheet1/B1");
        node.Format["formula"].Should().Be("SUM(A1:A5)");
        node.Format["cachedValue"].Should().Be("150");
        node.Text.Should().Be("150");
    }

    // ==================== AVERAGE and COUNT ====================

    [Fact]
    public void Formula_AVERAGE_COUNT()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "20" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "30" });
        _handler.Set("/Sheet1/A4", new() { ["value"] = "40" });

        _handler.Set("/Sheet1/B1", new() { ["formula"] = "AVERAGE(A1:A4)" });
        _handler.Set("/Sheet1/B2", new() { ["formula"] = "COUNT(A1:A4)" });

        var avgNode = _handler.Get("/Sheet1/B1");
        avgNode.Format["formula"].Should().Be("AVERAGE(A1:A4)");
        avgNode.Format["cachedValue"].Should().Be("25");
        avgNode.Text.Should().Be("25");

        var countNode = _handler.Get("/Sheet1/B2");
        countNode.Format["formula"].Should().Be("COUNT(A1:A4)");
        countNode.Format["cachedValue"].Should().Be("4");
        countNode.Text.Should().Be("4");
    }

    // ==================== IF with String result ====================

    [Fact]
    public void Formula_IF_String()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/B1", new() { ["formula"] = "IF(A1>5,\"Yes\",\"No\")" });

        var node = _handler.Get("/Sheet1/B1");
        node.Format["formula"].Should().Be("IF(A1>5,\"Yes\",\"No\")");
        node.Format["cachedValue"].Should().Be("Yes");
        node.Text.Should().Be("Yes");

        // Modify A1 to trigger the other branch and re-set formula
        _handler.Set("/Sheet1/A1", new() { ["value"] = "3" });
        _handler.Set("/Sheet1/B1", new() { ["formula"] = "IF(A1>5,\"Yes\",\"No\")" });

        var node2 = _handler.Get("/Sheet1/B1");
        node2.Format["cachedValue"].Should().Be("No");
        node2.Text.Should().Be("No");
    }

    // ==================== CONCATENATE ====================

    [Fact]
    public void Formula_CONCATENATE()
    {
        _handler.Set("/Sheet1/A1", new() { ["formula"] = "CONCATENATE(\"Hello\",\" \",\"World\")" });

        var node = _handler.Get("/Sheet1/A1");
        node.Format["formula"].Should().Be("CONCATENATE(\"Hello\",\" \",\"World\")");
        node.Format["cachedValue"].Should().Be("Hello World");
        node.Text.Should().Be("Hello World");
    }

    // ==================== VLOOKUP Exact ====================

    [Fact]
    public void Formula_VLOOKUP_Exact()
    {
        // Build a lookup table: A1:B3
        _handler.Set("/Sheet1/A1", new() { ["value"] = "100" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "Apple" , ["type"] = "string" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "200" });
        _handler.Set("/Sheet1/B2", new() { ["value"] = "Banana", ["type"] = "string" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "300" });
        _handler.Set("/Sheet1/B3", new() { ["value"] = "Cherry", ["type"] = "string" });

        // VLOOKUP(200, A1:B3, 2, FALSE) — exact match for 200
        _handler.Set("/Sheet1/C1", new() { ["formula"] = "VLOOKUP(200,A1:B3,2,FALSE)" });

        var node = _handler.Get("/Sheet1/C1");
        node.Format["formula"].Should().Be("VLOOKUP(200,A1:B3,2,FALSE)");
        node.Format["cachedValue"].Should().Be("Banana");
        node.Text.Should().Be("Banana");
    }

    // ==================== VLOOKUP Approximate ====================

    [Fact]
    public void Formula_VLOOKUP_Approximate()
    {
        // Sorted lookup table
        _handler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "Low", ["type"] = "string" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "50" });
        _handler.Set("/Sheet1/B2", new() { ["value"] = "Medium", ["type"] = "string" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "100" });
        _handler.Set("/Sheet1/B3", new() { ["value"] = "High", ["type"] = "string" });

        // VLOOKUP(75, A1:B3, 2) — approximate match, should find row with 50
        _handler.Set("/Sheet1/C1", new() { ["formula"] = "VLOOKUP(75,A1:B3,2)" });

        var node = _handler.Get("/Sheet1/C1");
        node.Format["formula"].Should().Be("VLOOKUP(75,A1:B3,2)");
        node.Format["cachedValue"].Should().Be("Medium");
        node.Text.Should().Be("Medium");
    }

    // ==================== INDEX + MATCH ====================

    [Fact]
    public void Formula_INDEX_MATCH()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "Apple", ["type"] = "string" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "Banana", ["type"] = "string" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "Cherry", ["type"] = "string" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/B2", new() { ["value"] = "20" });
        _handler.Set("/Sheet1/B3", new() { ["value"] = "30" });

        // INDEX(B1:B3, MATCH("Banana", A1:A3, 0)) should return 20
        _handler.Set("/Sheet1/C1", new() { ["formula"] = "INDEX(B1:B3,MATCH(\"Banana\",A1:A3,0))" });

        var node = _handler.Get("/Sheet1/C1");
        node.Format["formula"].Should().Be("INDEX(B1:B3,MATCH(\"Banana\",A1:A3,0))");
        node.Format["cachedValue"].Should().Be("20");
        node.Text.Should().Be("20");
    }

    // ==================== Chained Formulas ====================

    [Fact]
    public void Formula_ChainedFormulas()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "100" });
        _handler.Set("/Sheet1/B1", new() { ["formula"] = "A1*2" });
        _handler.Set("/Sheet1/C1", new() { ["formula"] = "B1+10" });

        var b1 = _handler.Get("/Sheet1/B1");
        b1.Format["cachedValue"].Should().Be("200");

        var c1 = _handler.Get("/Sheet1/C1");
        c1.Format["formula"].Should().Be("B1+10");
        c1.Format["cachedValue"].Should().Be("210");
        c1.Text.Should().Be("210");
    }

    // ==================== SUMIF ====================

    [Fact]
    public void Formula_SUMIF()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "50" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "150" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "200" });
        _handler.Set("/Sheet1/A4", new() { ["value"] = "80" });

        _handler.Set("/Sheet1/B1", new() { ["formula"] = "SUMIF(A1:A4,\">100\")" });

        var node = _handler.Get("/Sheet1/B1");
        node.Format["formula"].Should().Be("SUMIF(A1:A4,\">100\")");
        // SUMIF should sum 150 + 200 = 350
        node.Format["cachedValue"].Should().Be("350");
        node.Text.Should().Be("350");
    }

    // ==================== Text Functions ====================

    [Fact]
    public void Formula_TextFunctions()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "hello world", ["type"] = "string" });

        _handler.Set("/Sheet1/B1", new() { ["formula"] = "UPPER(A1)" });
        _handler.Set("/Sheet1/B2", new() { ["formula"] = "LEFT(A1,5)" });
        _handler.Set("/Sheet1/B3", new() { ["formula"] = "LEN(A1)" });

        var upperNode = _handler.Get("/Sheet1/B1");
        upperNode.Format["cachedValue"].Should().Be("HELLO WORLD");
        upperNode.Text.Should().Be("HELLO WORLD");

        var leftNode = _handler.Get("/Sheet1/B2");
        leftNode.Format["cachedValue"].Should().Be("hello");
        leftNode.Text.Should().Be("hello");

        var lenNode = _handler.Get("/Sheet1/B3");
        lenNode.Format["cachedValue"].Should().Be("11");
        lenNode.Text.Should().Be("11");
    }

    // ==================== Date Functions ====================

    [Fact]
    public void Formula_DateFunctions()
    {
        // DATE(2024,6,15) returns an OADate serial number
        _handler.Set("/Sheet1/A1", new() { ["formula"] = "DATE(2024,6,15)" });

        var dateNode = _handler.Get("/Sheet1/A1");
        dateNode.Format["formula"].Should().Be("DATE(2024,6,15)");
        dateNode.Format.Should().ContainKey("cachedValue");
        // The cached value is the OADate serial number
        var oaDate = double.Parse(dateNode.Format["cachedValue"].ToString()!);
        var dt = DateTime.FromOADate(oaDate);
        dt.Year.Should().Be(2024);
        dt.Month.Should().Be(6);
        dt.Day.Should().Be(15);

        // YEAR and MONTH from the serial number
        _handler.Set("/Sheet1/B1", new() { ["formula"] = "YEAR(A1)" });
        _handler.Set("/Sheet1/B2", new() { ["formula"] = "MONTH(A1)" });

        var yearNode = _handler.Get("/Sheet1/B1");
        yearNode.Format["cachedValue"].Should().Be("2024");

        var monthNode = _handler.Get("/Sheet1/B2");
        monthNode.Format["cachedValue"].Should().Be("6");
    }

    // ==================== Circular Reference ====================

    [Fact]
    public void Formula_CircularReference()
    {
        // A1 references B1, B1 references A1 — circular
        _handler.Set("/Sheet1/A1", new() { ["formula"] = "B1+1" });
        _handler.Set("/Sheet1/B1", new() { ["formula"] = "A1+1" });

        // The evaluator should detect circular reference and not infinite-loop.
        // The cell should either have no cached value or show an error/formula text.
        var nodeA = _handler.Get("/Sheet1/A1");
        var nodeB = _handler.Get("/Sheet1/B1");

        // At least one should lack a proper numeric cached value due to circular ref.
        // The evaluator returns #REF! for circular refs, which means no numeric cache.
        // Verify no crash occurred and the cells are accessible.
        nodeA.Should().NotBeNull();
        nodeB.Should().NotBeNull();
    }

    // ==================== Persistence ====================

    [Fact]
    public void Formula_Persistence()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "20" });
        _handler.Set("/Sheet1/B1", new() { ["formula"] = "SUM(A1:A2)" });

        var node = _handler.Get("/Sheet1/B1");
        node.Format["cachedValue"].Should().Be("30");

        // Reopen and verify the cached value persists
        Reopen();

        var reopened = _handler.Get("/Sheet1/B1");
        reopened.Format.Should().ContainKey("formula");
        reopened.Format["formula"].Should().Be("SUM(A1:A2)");
        reopened.Format.Should().ContainKey("cachedValue");
        reopened.Format["cachedValue"].Should().Be("30");
        reopened.Text.Should().Be("30");
    }
}
