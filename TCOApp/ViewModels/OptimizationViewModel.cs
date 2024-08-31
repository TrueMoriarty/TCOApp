using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using TCOApp.Models;
using TCOApp.Models.OutputModels;

namespace TCOApp.ViewModels;

/// <summary>
/// Класс, реализующий взаимодействие между моделями и отображениями в окне с оптимизацией.
/// </summary>
internal class OptimizationViewModel : INotifyPropertyChanged
{
    private string _pathTextBox;

    private List<VeneerResultModel> _veneerResultModel;
    private List<KRCResultModel> _krcResultModels;
    private List<AllCostModel> _allCostModels;

    private RelayCommand _choiceTheDirectory;
    private RelayCommand _saveFile;

    public string PathTextBox
    {
        get => _pathTextBox;
        set 
        {
            _pathTextBox = value;
            NotifyPropertyChanged(nameof(PathTextBox));
        }
    }

    public List<VeneerResultModel> VeneerResultModels
    {
        get => _veneerResultModel;
        set
        {
            _veneerResultModel = value;
            NotifyPropertyChanged(nameof(VeneerResultModels));
        }
    }
    public List<KRCResultModel> KRCResultModels
    {
        get => _krcResultModels;
        set
        {
            _krcResultModels = value;
            NotifyPropertyChanged(nameof(KRCResultModels));
        }
    }
    public List<AllCostModel> AllCostModels
    {
        get => _allCostModels;
        set
        {
            _allCostModels = value;
            NotifyPropertyChanged(nameof(AllCostModels));
        }
    }



    public RelayCommand ChoiceTheDirectory
    {
        get
        {
            return _choiceTheDirectory ?? (_choiceTheDirectory = new RelayCommand(obj =>
            {
                PathTextBox = OutputHandler.GetThePath();
                if (PathTextBox == null)
                {
                    MessageBox.Show("Directory doesn't exist");
                }
            }));
        }
    }
    public RelayCommand SaveFile
    {
        get
        {
            return _saveFile ?? (_saveFile = new RelayCommand(obj =>
            {
                OutputHandler.SaveInExcel(PathTextBox, KRCResultModels, AllCostModels);
            }));
        }
    }

    public OptimizationViewModel(List<VeneerResultModel> veneerResultModels, List<KRCResultModel> krcResultModels, List<AllCostModel> allCostModels)
    {
        VeneerResultModels = veneerResultModels;
        KRCResultModels = krcResultModels;
        AllCostModels = allCostModels;
    }

    // Шаблонный код для удовлетворения INotifyPropertyChanged
    public event PropertyChangedEventHandler PropertyChanged;
    public void NotifyPropertyChanged(string propertyName)
    {
        if (PropertyChanged is null) return;
        PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
    }
}
