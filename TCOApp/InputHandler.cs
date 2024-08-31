using ClosedXML.Excel;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Linq;
using TCOApp.Models;

namespace TCOApp;

/// <summary>
/// Реализация считывания данных из Excel
/// </summary>
internal class InputHandler
{

    public static string GetThePath()
    {
        // Create OpenFileDialog 
        OpenFileDialog dlg = new OpenFileDialog();

        // Set filter for file extension and default file extension 
        dlg.Multiselect = true;
        dlg.DefaultExt = ".xlsx";
        dlg.Filter = "Excel Files (*.xlsx;*.xls,*.xlsb,*.xlsm)|*.xlsx;*.xls,*.xlsb,*.xlsm";
        dlg.FilterIndex = 1;

        // Display OpenFileDialog by calling ShowDialog method 
        bool? result = dlg.ShowDialog();

        // Get the selected file name and display in a TextBox
        if (result == true)
        {
            // return path
            return dlg.FileName;
        }

        return null;
    }

    public static List<IXLWorksheet> GetWorksheetsFromPath(string path)
    {
        XLWorkbook workbook = new XLWorkbook(path);
        List<IXLWorksheet> worksheets = new List<IXLWorksheet>();

        foreach (IXLWorksheet worksheet in workbook.Worksheets)
        {
            worksheets.Add(worksheet);
        }

        return worksheets;
    }

    public static List<VeneerModel> GetDataFromVeneerTable(List<IXLWorksheet> worksheets)
    {
        List<VeneerModel> veneerModels = new List<VeneerModel>();

        worksheets[0].LastRowUsed().Delete();
        IEnumerable<IXLRangeRow> rows = worksheets[0].RangeUsed().RowsUsed().Skip(1);

        int k = 1;
        foreach (var row in rows)
        {
            foreach (var cell in row.Cells())
            {
                if (cell.IsEmpty())
                {
                    cell.Value = 0;
                }
            }
            veneerModels.Add(new VeneerModel()
            {
                Sort = row.Cell(k).Value.ToString(),
                Format = row.Cell(k + 1).Value.ToString(),
                Remainder = double.Parse(row.Cell(k + 2).Value.ToString()),
                Need = double.Parse(row.Cell(k + 3).Value.ToString())
            });
        }

        return veneerModels;
    }

    public static List<VeneerModel> GetDataFromVeneerTableAll(List<IXLWorksheet> worksheets)
    {
        List<VeneerModel> veneerModels = new List<VeneerModel>();

        worksheets[0].LastRowUsed().Delete();
        IEnumerable<IXLRangeRow> rows = worksheets[0].RangeUsed().RowsUsed().Skip(1);

        int k = 1;
        foreach (var row in rows)
        {
            foreach (var cell in row.Cells())
            {
                if (cell.IsEmpty())
                {
                    cell.Value = 0;
                }
            }
            veneerModels.Add(new VeneerModel()
            {
                Sort = row.Cell(k).Value.ToString(),
                Format = row.Cell(k + 1).Value.ToString(),
                Remainder = double.Parse(row.Cell(k + 2).Value.ToString()),
                Need = double.Parse(row.Cell(k + 3).Value.ToString())
            });
        }

        for (int i = 0; i < veneerModels.Count; ++i)
        {
            for (int j = 0; j < veneerModels.Count; ++j)
            {
                if (veneerModels[i].Sort == veneerModels[j].Sort && j != i)
                {
                    veneerModels[i].Need += veneerModels[j].Need;
                    veneerModels.Remove(veneerModels[j]);
                    j--;
                }
            }
        }

        List<VeneerModel> veneerModelsResult = new List<VeneerModel>();
        for (int i = 0; i < veneerModels.Count; ++i)
        {
            if (veneerModels[i].Sort != "увлаж" && veneerModels[i].Sort != "WG")
            {
                veneerModelsResult.Add(veneerModels[i]);
            }
            if (veneerModels[i].Sort == "WG")
            {
                for (int j = 0; j < veneerModelsResult.Count; ++j)
                {
                    if (veneerModelsResult[j].Sort == "CP")
                    {
                        veneerModelsResult[j].Need += veneerModels[i].Need;
                    }
                }
            }
        }

        return veneerModelsResult;
    }



    public static List<KRCModel> GetDataFromKRCTable(List<IXLWorksheet> worksheets)
    {
        List<KRCModel> krcModels = new List<KRCModel>();

        IEnumerable<IXLRangeRow> rows = worksheets[2].RangeUsed().RowsUsed().Skip(1);

        int i = 1;
        foreach (var row in rows)
        {
            foreach (var cell in row.Cells())
            {
                if (cell.IsEmpty())
                {
                    cell.Value = 1;
                }
            }
            krcModels.Add(new KRCModel()
            {
                Length = double.Parse(row.Cell(i).Value.ToString()),
                Sort = row.Cell(i + 1).Value.ToString(),
                LossesKrChur = double.Parse(row.Cell(i + 2).Value.ToString()),
                LossesChurSir = double.Parse(row.Cell(i + 3).Value.ToString()),
                LossesSirSuh = double.Parse(row.Cell(i + 4).Value.ToString()),
                LossesSuhObl = double.Parse(row.Cell(i + 5).Value.ToString()),
                PercentageOfWasteRS = double.Parse(row.Cell(i + 6).Value.ToString()),
                B = double.Parse(row.Cell(i + 7).Value.ToString()),
                S = double.Parse(row.Cell(i + 8).Value.ToString()),
                BB = double.Parse(row.Cell(i + 9).Value.ToString()),
                CP = double.Parse(row.Cell(i + 10).Value.ToString()),
                C = double.Parse(row.Cell(i + 11).Value.ToString()),
                One = double.Parse(row.Cell(i + 12).Value.ToString()),
                Two = double.Parse(row.Cell(i + 13).Value.ToString()),
                Three = double.Parse(row.Cell(i + 14).Value.ToString()),
                NF = double.Parse(row.Cell(i + 15).Value.ToString()),
                VeneerFromRidge = new double[]
                {
                    double.Parse(row.Cell(i + 7).Value.ToString()),
                    double.Parse(row.Cell(i + 8).Value.ToString()),
                    double.Parse(row.Cell(i + 9).Value.ToString()),
                    double.Parse(row.Cell(i + 10).Value.ToString()),
                    double.Parse(row.Cell(i + 11).Value.ToString()),
                    double.Parse(row.Cell(i + 12).Value.ToString()),
                    double.Parse(row.Cell(i + 13).Value.ToString()),
                    double.Parse(row.Cell(i + 14).Value.ToString()),
                    double.Parse(row.Cell(i + 15).Value.ToString()),
                }
            });
        }

        return krcModels;
    }

    public static List<RawMaterialsMarketModel> GetDataFromRawMaterialsMarketTable(List<IXLWorksheet> worksheets)
    {
        List<RawMaterialsMarketModel> rawMaterialsMarketModels = new List<RawMaterialsMarketModel>();

        IEnumerable<IXLRangeRow> rows = worksheets[1].RangeUsed().RowsUsed().Skip(1);

        int i = 1;
        foreach (var row in rows)
        {
            foreach (var cell in row.Cells())
            {
                if (cell.IsEmpty())
                {
                    cell.Value = 0;
                }
            }
            rawMaterialsMarketModels.Add(new RawMaterialsMarketModel()
            {
                Length = double.Parse(row.Cell(i).Value.ToString()),
                Sort = row.Cell(i + 1).Value.ToString(),
                Price = double.Parse(row.Cell(i + 2).Value.ToString()),
                Availability = double.Parse(row.Cell(i + 3).Value.ToString())
            });
        }
        return rawMaterialsMarketModels;
    }

    public static List<CoefficientsModel> GetDataFromCoefficientTable(List<IXLWorksheet> worksheets)
    {
        List<CoefficientsModel> coefficientModel = new List<CoefficientsModel>();
        IEnumerable<IXLRangeRow> rows = worksheets[4].RangeUsed().RowsUsed().Skip(1);

        var sortRow = rows.First();

        rows = rows.Skip(1);

        int i = 2;
        foreach (var row in rows)
        {
            foreach (var cell in row.Cells())
            {
                if (cell.Value.ToString() == "-")
                {
                    cell.Value = 0;
                }
            }

            coefficientModel.Add(new CoefficientsModel()
            {
                SortFrom = row.Cell(i).Value.ToString(),
                SortToFirst = sortRow.Cell(i + 1).Value.ToString(),
                SortToSecond = sortRow.Cell(i + 2).Value.ToString(),
                SortToThird = sortRow.Cell(i + 3).Value.ToString(),
                SortToFourth = sortRow.Cell(i + 4).Value.ToString(),
                SortToFive = sortRow.Cell(i + 5).Value.ToString(),
                SortToSixth = sortRow.Cell(i + 6).Value.ToString(),
                FirstCoefficient = double.Parse(row.Cell(i + 1).Value.ToString()),
                SecondCoefficient = double.Parse(row.Cell(i + 2).Value.ToString()),
                ThirdCoefficient = double.Parse(row.Cell(i + 3).Value.ToString()),
                FourthCoefficient = double.Parse(row.Cell(i + 4).Value.ToString()),
                FiveCoefficient = double.Parse(row.Cell(i + 5).Value.ToString()),
                SixthCoefficient = double.Parse(row.Cell(i + 6).Value.ToString()),
            });
        }
        return coefficientModel;
    }

    // методы для тестов
    public static List<VeneerModel> GetDataFromVeneerFiveFive(List<VeneerModel> veneerModels)
    {
        List<VeneerModel> veneerModelsReturn = new List<VeneerModel>();
        int j = 0;
        for (int i = 0; i < veneerModels.Count; ++i)
        {
            if (veneerModels[i].Format == "5х5" && veneerModels[i].Sort != "увлаж")
            {
                if (veneerModels[i].Sort == "WG")
                {
                    veneerModelsReturn[j - 1].Need += veneerModels[i].Need;
                }
                if (veneerModels[i].Sort != "WG")
                {
                    veneerModelsReturn.Add(veneerModels[i]);
                    j++;
                }
            }
        }
        return veneerModelsReturn;
    }

    public static List<VeneerModel> GetDataFromVeneerFiveTen(List<VeneerModel> veneerModels)
    {
        List<VeneerModel> veneerModelsReturn = new List<VeneerModel>();
        int j = 0;
        for (int i = 0; i < veneerModels.Count; ++i)
        {
            if (veneerModels[i].Format == "5х10" && veneerModels[i].Sort != "увлаж")
            {
                if (veneerModels[i].Sort == "WG")
                {
                    veneerModelsReturn[j - 1].Need += veneerModels[i].Need;
                }
                if (veneerModels[i].Sort != "WG")
                {
                    veneerModelsReturn.Add(veneerModels[i]);
                    j++;
                }
            }
        }
        return veneerModelsReturn;
    }

    public static List<VeneerModel> GetDataFromVeneerTenFive(List<VeneerModel> veneerModels)
    {
        List<VeneerModel> veneerModelsReturn = new List<VeneerModel>();
        int j = 0;
        for (int i = 0; i < veneerModels.Count; ++i)
        {
            if (veneerModels[i].Format == "10х5" && veneerModels[i].Sort != "увлаж")
            {
                if (veneerModels[i].Sort == "WG")
                {
                    veneerModelsReturn[j - 1].Need += veneerModels[i].Need;
                }
                if (veneerModels[i].Sort != "WG")
                {
                    veneerModelsReturn.Add(veneerModels[i]);
                    j++;
                }
            }
        }
        return veneerModelsReturn;
    }

    public static List<VeneerModel> GetDataFromVeneerFourEight(List<VeneerModel> veneerModels)
    {
        List<VeneerModel> veneerModelsReturn = new List<VeneerModel>();
        int j = 0;
        for (int i = 0; i < veneerModels.Count; ++i)
        {
            if (veneerModels[i].Format == "4х8" && veneerModels[i].Sort != "увлаж")
            {
                if (veneerModels[i].Sort == "WG")
                {
                    veneerModelsReturn[j - 1].Need += veneerModels[i].Need;
                }
                if (veneerModels[i].Sort != "WG")
                {
                    veneerModelsReturn.Add(veneerModels[i]);
                    j++;
                }
            }
        }
        return veneerModelsReturn;
    }

    public static List<VeneerModel> GetDataFromVeneerFourFour(List<VeneerModel> veneerModels)
    {
        List<VeneerModel> veneerModelsReturn = new List<VeneerModel>();
        int j = 0;
        for (int i = 0; i < veneerModels.Count; ++i)
        {
            if (veneerModels[i].Format == "4х4" && veneerModels[i].Sort != "увлаж")
            {
                if (veneerModels[i].Sort == "WG")
                {
                    veneerModelsReturn[j - 1].Need += veneerModels[i].Need;
                }
                if (veneerModels[i].Sort != "WG")
                {
                    veneerModelsReturn.Add(veneerModels[i]);
                    j++;
                }
            }
        }
        return veneerModelsReturn;
    }

    public static List<VeneerModel> GetDataFromVeneerEightFour(List<VeneerModel> veneerModels)
    {
        List<VeneerModel> veneerModelsReturn = new List<VeneerModel>();
        int j = 0;
        for (int i = 0; i < veneerModels.Count; ++i)
        {
            if (veneerModels[i].Format == "8х4" && veneerModels[i].Sort != "увлаж")
            {
                if (veneerModels[i].Sort == "WG")
                {
                    veneerModelsReturn[j - 1].Need += veneerModels[i].Need;
                }
                if (veneerModels[i].Sort != "WG")
                {
                    veneerModelsReturn.Add(veneerModels[i]);
                    j++;
                }
            }
        }
        return veneerModelsReturn;
    }

    /*public static List<EquipmentModel> GetDataFromEquipmentTable(List<IXLWorksheet> worksheets)
    {
        List<EquipmentModel> equipmentModels = new List<EquipmentModel>();

        IEnumerable<IXLRangeRow> rows = worksheets[3].RangeUsed().RowsUsed().Skip(1);

        int i = 1;
        foreach (var row in rows)
        {
            foreach (var cell in row.Cells())
            {
                if (cell.IsEmpty())
                {
                    cell.Value = 0;
                }
            }
            equipmentModels.Add(new EquipmentModel()
            {
                Stage = row.Cell(i).Value.ToString(),
                Diameter = double.Parse(row.Cell(i + 1).Value.ToString()),
                Bucking  double.Parse(row.Cell(i+2).Value.ToString()),

            });
        }
        return equipmentModels;
    }*/
}
