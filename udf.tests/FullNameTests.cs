using FluentAssertions;
using gmafffff.excel.udf.ФИО;

namespace udf.tests;

public class ФИОТесты {
    public static IEnumerable<object[]> ПримерыФио => new[] {
        new object[] { "Гришкин Максим Александрович", false, "Гришкин М.А." },
        new object[] { "Гришкин М.А.", false, "Гришкин М.А." },
        new object[] { "Гришкин Максим Александрович", true, "М.А. Гришкин" }
    };

    [Theory]
    [MemberData(nameof(ПримерыФио))]
    public void ВозможныйФормат(string полн, bool слева, string сокр) {
        ФИО.СократитьФио(полн, слева).Should().Be(сокр);
    }
}