
namespace TCOApp.Models;

/// <summary>
/// Класс описывающий таблицу c шпоном, где каждое поле столбец в таблице
/// </summary>
public class VeneerModel
{
    public string Sort { get; set; }
    public string Format { get; set; }
    public double Remainder { get; set; }
    public double Need { get; set; }
}
