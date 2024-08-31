
namespace TCOApp.Models;

/// <summary>
/// Класс, описывающий результат работы модели
/// </summary>
public class VeneerResultModel
{
    public string VeneerSort { get; set; }
    public string RidgeSort { get; set; }
    public double Result { get; set; }
    public double Need { get; set; }
    public double Get { get; set; }
}
