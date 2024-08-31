namespace TCOApp.Models.OutputModels;

public class AllCostModel
{
    public string RidgeSort { get; set; }
    public double RidgeCount { get; set; }
    public double RidgeCost { get; set; }
    public double UntranslatedVeneerCost { get; set; }
    public double TranslatedVeneerCost { get; set; }
    public double VeneerCostWithoutResidues { get; set; }
    public double VeneerCostWithResidues { get; set; }
}
