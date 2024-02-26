using AutoFixture;
using AutoFixture.AutoMoq;
using FluentAssertions;
using gmafffff.excel.udf.Excel.Команды;
using gmafffff.excel.udf.Excel.Сетка;
using Moq;

namespace udf.tests.Excel.Команды;

public class ИзмениВидимостьРядаКомандаТесты {
    private readonly Fixture fixture = new();

    public ИзмениВидимостьРядаКомандаТесты() {
        fixture.Customize(new AutoMoqCustomization { ConfigureMembers = true });
        fixture.Inject(new IntPtr(0));
        fixture.Behaviors.OfType<ThrowingRecursionBehavior>().ToList()
               .ForEach(b => fixture.Behaviors.Remove(b));
        fixture.Behaviors.Add(new OmitOnRecursionBehavior());
    }

    [Fact]
    public void МожноУпаковатьОднотипныеКоманды() {
        // Подготовка
        var ряд = fixture.Freeze<Mock<РядСетки>>();
        ряд.Setup(стр => стр.ВидимЛи()).Returns(true);

        var ряды = fixture.CreateMany<РядСетки>(10).DistinctBy(р => р.Номер);

        var количество = ряды.Count();
        var ряды_дубл  = ряды.Concat(ряды); // Дубликаты

        var авторазмер = true;
        var показать   = false;
        var размер     = 1;

        var команды_вх =
            from р in ряды_дубл
            select new ИзмениВидимостьРядаКоманда(р, показать, размер, авторазмер);

        // Действие
        var команды_упак = ИзмениВидимостьРядаКоманда.Упаковать(команды_вх);

        // Проверка
        команды_упак.Should().ContainSingle();
        var команда_у = команды_упак.Single();

        команда_у.Авторазмер.Should().Be(авторазмер);
        команда_у.Показать.Should().Be(показать);
        команда_у.Размер.Should().Be(размер);
        команда_у.Ряды.Should().Equal(ряды);
    }
}