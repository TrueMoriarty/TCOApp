using System.Collections.Generic;
using System.Windows;
using TCOApp.Models;
using TCOApp.Models.OutputModels;
using TCOApp.ViewModels;

namespace TCOApp
{
    /// <summary>
    /// Логика взаимодействия для OptimizationView.xaml
    /// </summary>
    public partial class OptimizationView : Window
    {
        public OptimizationView(List<VeneerResultModel> veneerResultModels, List<KRCResultModel> krcResultModels,
            List<AllCostModel> allCostModels)
        {
            InitializeComponent();
            DataContext = new OptimizationViewModel(veneerResultModels, krcResultModels, allCostModels);
        }
    }
}
