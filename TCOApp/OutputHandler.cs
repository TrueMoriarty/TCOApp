using ClosedXML.Excel;
using Google.OrTools.LinearSolver;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using TCOApp.Models;
using TCOApp.Models.OutputModels;

namespace TCOApp;

/// <summary>
/// Реализация вывода работы модели в Excel
/// </summary>
class OutputHandler
{
    public static string GetThePath()
    {
        FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

        // Show the FolderBrowserDialog.
        DialogResult result = folderBrowserDialog.ShowDialog();
        if (result == DialogResult.OK)
        {
            return folderBrowserDialog.SelectedPath;
        }
        return null;

    }

    public static void SaveInExcel(string path, List<KRCResultModel> krcResultModels, List<AllCostModel> allCostModels)
    {
        XLWorkbook workbook = new XLWorkbook();
        IXLWorksheet worksheet = workbook.AddWorksheet("OptimizedTable");

        for (int i = 0; i < krcResultModels.Count; ++i)
        {
            worksheet.Cell(i + 1, 1).Value = krcResultModels[i].Sort;
            worksheet.Cell(i + 1, 2).Value = krcResultModels[i].Amount;
            worksheet.Cell(i + 1, 3).Value = krcResultModels[i].B;
            worksheet.Cell(i + 1, 4).Value = krcResultModels[i].S;
            worksheet.Cell(i + 1, 5).Value = krcResultModels[i].BB;
            worksheet.Cell(i + 1, 6).Value = krcResultModels[i].CP;
            worksheet.Cell(i + 1, 7).Value = krcResultModels[i].C;
            worksheet.Cell(i + 1, 8).Value = krcResultModels[i].One;
            worksheet.Cell(i + 1, 9).Value = krcResultModels[i].Two;
            worksheet.Cell(i + 1, 10).Value = krcResultModels[i].Three;
            worksheet.Cell(i + 1, 11).Value = krcResultModels[i].NF;
            worksheet.Cell(i + 1, 12).Value = allCostModels[i].RidgeCost;

        }

        workbook.SaveAs(path + "\\Result2_1.xlsx");
    }

    public static List<VeneerResultModel> MapDataFromSolverToVeneerModel(List<Variable> variables, List<VeneerModel> veneerModels, List<KRCModel> krcModels)
    {
        List<VeneerResultModel> veneerResultModels = new List<VeneerResultModel>();

        for (int i = 0; i < veneerModels.Count; ++i)
        {
            for (int j = 0; j < krcModels.Count; ++j)
            {
                veneerResultModels.Add(new VeneerResultModel()
                {
                    VeneerSort = veneerModels[i].Sort,
                    RidgeSort = krcModels[j].Sort,
                    Result = Math.Ceiling(variables[j].SolutionValue()),
                    Get = Math.Ceiling(variables[j].SolutionValue() * 1 /
                        (krcModels[j].LossesKrChur * krcModels[j].LossesChurSir *
                        krcModels[j].LossesSirSuh * krcModels[j].LossesSuhObl) *
                        krcModels[j].VeneerFromRidge[i] / 100),
                    Need = veneerModels[i].Need
                });
            }
        }

        return veneerResultModels;
    }

    public static List<KRCResultModel> MapDataFromSolverToRidgeModel(List<Variable> variables, List<KRCModel> krcModels, List<VeneerModel> veneerModels)
    {
        List<KRCResultModel> ridgeModels = new List<KRCResultModel>();
        double[] resultVeneers = new double[veneerModels.Count];
        for (int i = 0; i < variables.Count; ++i)
        {
            ridgeModels.Add(new KRCResultModel()
            {
                Sort = krcModels[i].Sort,
                Amount = Math.Ceiling(variables[i].SolutionValue()),
            });
        }
        for (int i = 0; i < variables.Count; ++i)
        {
            for (int j = 0; j < 9; ++j)
            {
                resultVeneers[j] = Math.Ceiling(variables[i].SolutionValue() / (krcModels[i].LossesKrChur * krcModels[i].LossesChurSir *
                    krcModels[i].LossesSirSuh * krcModels[i].LossesSuhObl) *
                    krcModels[i].VeneerFromRidge[j] / 100);
            }
            ridgeModels[i].B = resultVeneers[0];
            ridgeModels[i].S = resultVeneers[1];
            ridgeModels[i].BB = resultVeneers[2];
            ridgeModels[i].CP = resultVeneers[3];
            ridgeModels[i].C = resultVeneers[4];
            ridgeModels[i].One = resultVeneers[5];
            ridgeModels[i].Two = resultVeneers[6];
            ridgeModels[i].Three = resultVeneers[7];
            ridgeModels[i].NF = resultVeneers[8];
        }
        return ridgeModels;
    }

    public static List<AllCostModel> MapDataFromSolverToAllCostModel(List<Variable> variables, List<KRCModel> krcModels, 
        List<RawMaterialsMarketModel> rawMaterialsMarketModels)
    {
        List<AllCostModel> allCostModels = new List<AllCostModel>();

        for (int i = 0; i < variables.Count; ++i)
        {
            allCostModels.Add(new AllCostModel()
            {
                RidgeSort = krcModels[i].Sort,
                RidgeCount = Math.Ceiling(variables[i].SolutionValue()),
                RidgeCost = Math.Ceiling(variables[i].SolutionValue() * rawMaterialsMarketModels[i].Price),
            });
        }

        return allCostModels;
    }
}
