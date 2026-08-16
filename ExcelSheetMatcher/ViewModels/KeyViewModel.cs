using System.Collections.Generic;
using System.Collections.ObjectModel;
using KeySettings = ExcelSheetMatcher.Models.KeySettings;

namespace ExcelSheetMatcher.ViewModels
{
    public class KeyViewModel: BaseVM
    {
        private ObservableCollection<KeySettings> _headers;

        public ObservableCollection<KeySettings> Headers
        {
            get => _headers;
            set
            {
                _headers = value;
                OnPropertyChanged();
            }
        }

        
    }
}
