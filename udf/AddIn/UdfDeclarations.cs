using ExcelDna.Integration;

namespace gmafffff.excel.udf.AddIn;

public static class ������� {
    private const string ������������ = "������� �� gmaFFFFF";

    [ExcelFunction(Name = "�������������", Category = ������������,
        Description = "���������� ����� � ������ ��������")]
    public static string �������������(
        [ExcelArgument(Name = "�����������", Description = "�����, ������� ���������� �������� ��������")]
        double �����,
        [ExcelArgument(Name = "������",
            Description =
                @"�[n](�����, n - ������ ����� �������), �/�[n][�[�]] (�����/������� �����, � - �������, � - � ���������, n - ������), �/�[�] (������ �������/�������, � - �����������). ������: ""�2 �� (��� � �2 �)""")]
        string ������ = "") {
        return ��������������.��������������.�������������(�����, ������);
    }

    [ExcelFunction(Name = "����������", Category = ������������,
        Description = "���������� �� ������ �� ���������� ������� �����")]
    public static double �������(
        [ExcelArgument(Name = "�����", Description = "����������� �����")]
        double �����,
        [ExcelArgument(Name = "������",
            Description = @"����� ������, �� ������� ���������� ����������.
���� ����� �������������, �� ���������� �� �������� ����� �������.
�������� 15 ������")]
        short ������) {
        return (������, Math.Pow(10, -������)) switch {
            (> 15, _) => Math.Round(�����, 15, MidpointRounding.ToEven),
            (>= 0, _) => Math.Round(�����, ������, MidpointRounding.ToEven),
            (< 0, var �������) => Math.Round(����� / �������, 0, MidpointRounding.ToEven) * �������
        };
    }
}