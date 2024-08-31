using ClosedXML.Excel;
using Google.OrTools.LinearSolver;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using TCOApp.Models;
using TCOApp.Models.OutputModels;

namespace TCOApp.ViewModels;

/// <summary>
/// Класс, реализующий взаимодействие между моделями и отображениями в окне с выбором Excel.
/// </summary>
internal class ViewModel : INotifyPropertyChanged
{
    private string _pathTextBox;
    private List<IXLWorksheet> _worksheets;

    private List<VeneerModel> _veneerModels;
    private List<VeneerModel> _veneerModelsFiveFive;
    private List<VeneerModel> _veneerModelsFiveTen;
    private List<VeneerModel> _veneerModelsTenFive;
    private List<VeneerModel> _veneerModelsFourEight;
    private List<VeneerModel> _veneerModelsFourFour;
    private List<VeneerModel> _veneerModelsEightFour;

    private List<KRCModel> _krcModels;
    private List<EquipmentModel> _equipmentModels;
    private List<RawMaterialsMarketModel> _rawMaterialsMarketModels;
    private List<CoefficientsModel> _coefficientsModels;

    private List<EquipmentPerformanceLimitationModel> _equipmentPerformanceLimitation;
    private List<NfVeneerLimitationModel> _nfVeneerLimitation;
    private List<RawMaterialLimitationModel> _rawMaterialLimitation;

    private List<Variable> _variables;

    private List<VeneerResultModel> _veneerResultModel;
    private List<KRCResultModel> _krcResultModels;
    private List<AllCostModel> _allCostModels;

    private RelayCommand _choiceThePath;
    private RelayCommand _printResult;
    private RelayCommand _optimize;

    public string PathTextBox
    {
        get => _pathTextBox;
        set
        {
            _pathTextBox = value;
            if (string.IsNullOrEmpty(_pathTextBox))
            {
                MessageBox.Show("The field must not be empty", "Alert", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            NotifyPropertyChanged(nameof(PathTextBox));
        }
    }
    public List<IXLWorksheet> Worksheets
    {
        get => _worksheets;
        set
        {
            _worksheets = value;
            if (_worksheets == null)
            {
                MessageBox.Show("This Excel file does not contain tables", "Alert", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            NotifyPropertyChanged(nameof(Worksheets));
        }
    }

    public List<VeneerModel> VeneerModels
    {
        get => _veneerModels;
        set
        {
            _veneerModels = value;
            NotifyPropertyChanged(nameof(VeneerModels));
        }
    }
    public List<VeneerModel> VeneerModelsFiveFive
    {
        get => _veneerModelsFiveFive;
        set
        {
            _veneerModelsFiveFive = value;
            NotifyPropertyChanged(nameof(VeneerModelsFiveFive));
        }
    }
    public List<VeneerModel> VeneerModelsFiveTen
    {
        get => _veneerModelsFiveTen;
        set
        {
            _veneerModelsFiveTen = value;
            NotifyPropertyChanged(nameof(VeneerModelsFiveTen));
        }
    }
    public List<VeneerModel> VeneerModelsTenFive
    {
        get => _veneerModelsTenFive;
        set
        {
            _veneerModelsTenFive = value;
            NotifyPropertyChanged(nameof(VeneerModelsTenFive));
        }
    }
    public List<VeneerModel> VeneerModelsFourEight
    {
        get => _veneerModelsFourEight;
        set
        {
            _veneerModelsFourEight = value;
            NotifyPropertyChanged(nameof(VeneerModelsFourEight));
        }
    }
    public List<VeneerModel> VeneerModelsFourFour
    {
        get => _veneerModelsFourFour;
        set
        {
            _veneerModelsFourFour = value;
            NotifyPropertyChanged(nameof(VeneerModelsFourFour));
        }
    }
    public List<VeneerModel> VeneerModelsEightFour
    {
        get => _veneerModelsEightFour;
        set
        {
            _veneerModelsEightFour = value;
            NotifyPropertyChanged(nameof(VeneerModelsEightFour));
        }
    }


    public List<KRCModel> KRCModels
    {
        get => _krcModels;
        set
        {
            _krcModels = value;
            NotifyPropertyChanged(nameof(KRCModels));
        }
    }
    public List<EquipmentModel> EquipmentModels
    {
        get => _equipmentModels;
        set
        {
            _equipmentModels = value;
            NotifyPropertyChanged(nameof(EquipmentModels));
        }
    }
    public List<RawMaterialsMarketModel> RawMaterialsMarketModels
    {
        get => _rawMaterialsMarketModels;
        set
        {
            _rawMaterialsMarketModels = value;
            NotifyPropertyChanged(nameof(RawMaterialsMarketModels));
        }
    }
    public List<CoefficientsModel> CoefficientsModels
    {
        get => _coefficientsModels;
        set
        {
            _coefficientsModels = value;
            NotifyPropertyChanged(nameof(CoefficientsModels));
        }
    }

    public List<Variable> Variables
    {
        get => _variables;
        set
        {
            _variables = value;
            NotifyPropertyChanged(nameof(Variables));
        }
    }

    public List<VeneerResultModel> VeneerResultModels
    {
        get => _veneerResultModel;
        set
        {
            _veneerResultModel = value;
        }
    }
    public List<KRCResultModel> KRCResultModels
    {
        get => _krcResultModels;
        set
        {
            _krcResultModels = value;
        }
    }
    public List<AllCostModel> AllCostModels
    {
        get => _allCostModels;
        set
        {
            _allCostModels = value;
        }
    }

    public List<EquipmentPerformanceLimitationModel> EquipmentPerformanceLimitation
    {
        get => _equipmentPerformanceLimitation;
        set
        {
            _equipmentPerformanceLimitation = value;
        }
    }
    public List<NfVeneerLimitationModel> NfVeneerLimitation
    {
        get => _nfVeneerLimitation;
        set
        {
            _nfVeneerLimitation = value;
        }
    }
    public List<RawMaterialLimitationModel> RawMaterialLimitation
    {
        get => _rawMaterialLimitation;
        set
        {
            _rawMaterialLimitation = value;
        }
    }

    public ViewModel()
    {
        EquipmentPerformanceLimitation = InitializationEquipmentPerformanceLimitation();
        NfVeneerLimitation = InitializationNfVeneerLimitation();
        RawMaterialLimitation = InitializationRawMaterialLimitation();
    } 

    private List<EquipmentPerformanceLimitationModel> InitializationEquipmentPerformanceLimitation()
    {
        string[] names = { "Линии раскряжевки", "Линии лущения", "Линии сушки",
            "Линии ребросклеивания", "Чурак", "Сухой шпон", "Облагороженный шпон"};

        List<EquipmentPerformanceLimitationModel> initialization = new List<EquipmentPerformanceLimitationModel>();
        foreach (string name in names) 
        {
            initialization.Add(new EquipmentPerformanceLimitationModel()
            {
                EquipmentName = name,
                IsChecked = false,
            });
        }
        return initialization;
    }
    private List<NfVeneerLimitationModel> InitializationNfVeneerLimitation()
    {

        string[] names = { "Объем н/ф меньше производительности р/б", "Объем н/ф больше производительности р/б"};

        List<NfVeneerLimitationModel> initialization = new List<NfVeneerLimitationModel>();
        foreach (string name in names)
        {
            initialization.Add(new NfVeneerLimitationModel()
            {
                Name = name,
                IsChecked = false,
            });
        }
        return initialization;
    }
    private List<RawMaterialLimitationModel> InitializationRawMaterialLimitation()
    {

        string[] names = { "Объем сырья меньше, чем указано в наличии" };

        List<RawMaterialLimitationModel> initialization = new List<RawMaterialLimitationModel>();
        foreach (string name in names)
        {
            initialization.Add(new RawMaterialLimitationModel()
            {
                Name = name,
                IsChecked = false,
            });
        }
        return initialization;
    }

    public RelayCommand ChoiceThePath
    {
        get
        {
            return _choiceThePath ??
                (_choiceThePath = new RelayCommand(obj =>
                {
                    PathTextBox = InputHandler.GetThePath();
                    Worksheets = InputHandler.GetWorksheetsFromPath(PathTextBox);
                }));
        }
    }
    public RelayCommand PrintResult
    {
        get
        {
            return _printResult ??
              (_printResult = new RelayCommand(obj =>
              {
                  VeneerModels = InputHandler.GetDataFromVeneerTable(Worksheets);
                  VeneerModelsFiveFive = InputHandler.GetDataFromVeneerFiveFive(VeneerModels);
                  VeneerModelsFiveTen = InputHandler.GetDataFromVeneerFiveTen(VeneerModels);
                  VeneerModelsTenFive = InputHandler.GetDataFromVeneerTenFive(VeneerModels);
                  VeneerModelsFourEight = InputHandler.GetDataFromVeneerFourEight(VeneerModels);
                  VeneerModelsFourFour = InputHandler.GetDataFromVeneerFourFour(VeneerModels);
                  VeneerModelsEightFour = InputHandler.GetDataFromVeneerEightFour(VeneerModels);

                  KRCModels = InputHandler.GetDataFromKRCTable(Worksheets);
                  //EquipmentModels = InputHandler.GetDataFromEquipmentTable(Worksheets);
                  RawMaterialsMarketModels = InputHandler.GetDataFromRawMaterialsMarketTable(Worksheets);
                  CoefficientsModels = InputHandler.GetDataFromCoefficientTable(Worksheets);
              }));
        }
    }
    public RelayCommand Optimize
    {
        get
        {
            return _optimize ??
                (_optimize = new RelayCommand(obj =>
                {

                    Variables = LinearOptimization.DoOptimization(VeneerModelsTenFive, KRCModels, RawMaterialsMarketModels);
                    VeneerResultModels = OutputHandler.MapDataFromSolverToVeneerModel(Variables, VeneerModelsTenFive, KRCModels);
                    KRCResultModels = OutputHandler.MapDataFromSolverToRidgeModel(Variables, KRCModels, VeneerModelsTenFive);
                    AllCostModels = OutputHandler.MapDataFromSolverToAllCostModel(Variables, KRCModels, RawMaterialsMarketModels);

                    OptimizationView optimizationView = new OptimizationView(VeneerResultModels, KRCResultModels, AllCostModels);
                    optimizationView.Show();
                }));
        }
    }

    // Шаблонный код для удовлетворения INotifyPropertyChanged
    public event PropertyChangedEventHandler PropertyChanged;
    public void NotifyPropertyChanged(string propertyName)
    {
        if (PropertyChanged is null) return;
        PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
    }
}