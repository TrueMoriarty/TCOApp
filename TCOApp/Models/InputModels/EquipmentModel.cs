using System.Collections.Generic;

namespace TCOApp.Models;

/// <summary>
/// Класс описывающий таблицу с линиями, где каждое поле столбец в таблице
/// </summary>
public class EquipmentModel
{
    public string Stage { get; set; }
    public double Diameter { get; set; }
    public Dictionary<string, double> Bucking { get; set; }
    public Dictionary<string, double> Peeling { get; set; }
    public Dictionary<string, double> Drying { get; set; }
    public Dictionary<string, double> EdgeGluing { get; set; }
    public Dictionary<string, double> Repair { get; set; }
    public Dictionary<string, double> Splice { get; set; }
}
