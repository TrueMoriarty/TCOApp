
namespace TCOApp.Models;

/// <summary>
/// Класс описывающий таблицу "сырье рынок", где каждое поле столбец в таблице
/// </summary>
public class RawMaterialsMarketModel
{
    public double Length { get; set; }
    public string Sort { get; set; }
    public double Price { get; set; }
    public double Availability { get; set; }
}
