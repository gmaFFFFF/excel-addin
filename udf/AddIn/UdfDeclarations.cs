using ExcelDna.Integration;

namespace gmafffff.excel.udf.AddIn;

public static class ������� {
    [ExcelFunction(Name = "�������������", Category = "������� �� gmaFFFFF",
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
}