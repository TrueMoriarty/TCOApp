
namespace TCOApp.Models;

/// <summary>
/// Класс описывающий таблицу "сырье выход", где каждое поле столбец в таблице
/// </summary>
public class KRCResultModel
{
    public string Sort { get; set; }
    public double Amount { get; set; }
    public double B { get; set; }
    public double S { get; set; }
    public double BB { get; set; }
    public double CP { get; set; }
    public double C { get; set; }
    public double One { get; set; }
    public double Two { get; set; }
    public double Three { get; set; }
    public double NF { get; set; }
}