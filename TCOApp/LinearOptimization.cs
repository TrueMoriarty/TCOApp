using System;
using System.Collections.Generic;
using Google.OrTools.LinearSolver;
using TCOApp.Models;

namespace TCOApp;

internal class LinearOptimization
{
    public static List<Variable> DoOptimization(List<VeneerModel> veneerModels, List<KRCModel> krcModels, List<RawMaterialsMarketModel> rawMaterialsMarketModels)
    {
        Solver solver = Solver.CreateSolver("GLOP");
        if (solver is null)
        {
            return null;
        }

        List<Variable> variables = CreateVariables(krcModels, solver);
        List<double> coefficients = new List<double>();

        for (int i = 0; i < veneerModels.Count; ++i)
        {
            var constraint = solver.MakeConstraint(veneerModels[i].Need, double.PositiveInfinity);

            for (int j = 0; j < krcModels.Count; ++j)
            {
                double coefficient = 1 / (krcModels[j].LossesKrChur *
                                          krcModels[j].LossesChurSir *
                                          krcModels[j].LossesSirSuh *
                                          krcModels[j].LossesSuhObl);
                coefficient *= krcModels[j].VeneerFromRidge[i] / 100;

                coefficients.Add(coefficient);

                // Добавляем вклад текущего сорта шпона
                constraint.SetCoefficient(variables[j], coefficient);

                // Добавляем вклад всех более высоких сортов шпона
                /*for (int k = 0; k < j; ++k)  // k - более высокие сорта
                {
                    constraint.SetCoefficient(variables[k], coefficient);
                }*/
            }
        }

        Objective objective = solver.Objective();
        for (int i = 0; i < variables.Count; ++i)
        {
            objective.SetCoefficient(variables[i], 1);
        }
        objective.SetMinimization();

        Solver.ResultStatus resultStatus = solver.Solve();
        if (resultStatus != Solver.ResultStatus.OPTIMAL)
        {
            Console.WriteLine("The problem does not have an optimal solution!");
            return null;
        }

        return variables;
    }

    private static List<Variable> CreateVariables(List<KRCModel> krcModels, Solver solver)
    {
        List<Variable> variables = new List<Variable>();
        for (int i = 0; i < krcModels.Count; ++i)
        {
            variables.Add(solver.MakeNumVar(0.0, double.PositiveInfinity, krcModels[i].Sort));
        }
        return variables;
    }
}
