
namespace TCOApp.Models;

/// <summary>
/// Класс описывающий таблицу "сырье выход", где каждое поле столбец в таблице
/// </summary>
public class KRCModel
{
    public double Length { get; set; }
    public string Sort { get; set; }
    public double LossesKrChur { get; set; }
    public double LossesChurSir { get; set; }
    public double LossesSirSuh { get; set; }
    public double LossesSuhObl { get; set; }
    public double PercentageOfWasteRS { get; set; }
    public double B { get; set; }
    public double S { get; set; }
    public double BB { get; set; }
    public double CP { get; set; }
    public double C { get; set; }
    public double One { get; set; }
    public double Two { get; set; }
    public double Three { get; set; }
    public double NF { get; set; }

    // Поле описывает виды шпона выше, в виде массива, нужен для работы с OR-Tools
    public double[] VeneerFromRidge { get; set; }
}