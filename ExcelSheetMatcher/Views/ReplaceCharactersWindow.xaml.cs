using System.Windows;
using ExcelSheetMatcher.ViewModels;

namespace ExcelSheetMatcher.Views
{
    public partial class ReplaceCharactersWindow : Window
    {
        private bool _isApplied = false;
        public bool IsApplied => _isApplied;
        public ReplaceCharactersWindow()
        {
            InitializeComponent();
            ((ReplaceCharactersViewModel)DataContext).OnApplyButton += () =>
            {
                _isApplied = true;
                this.Close();
            };
        }
    }
}
