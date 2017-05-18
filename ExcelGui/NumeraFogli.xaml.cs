using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using ClassiDiScambio.NumeraFogli;


namespace ExcelGui
{
    /// <summary>
    /// Logica di interazione per NumeraFogli.xaml
    /// </summary>
    public partial class NumeraFogli : Window
    {
        Sheet data;

        int _noOfErrorsOnScreen = 0;

        public NumeraFogli(Sheet data)
        {
            this.data = data;

            InitializeComponent();
            DataContext = this.data;

            ListaPagine.ItemsSource = this.data.contenuto;
        }

        private void Esegui_Click(object sender, RoutedEventArgs e)
        {
            data.Write();
            this.Close();
        }

        private void OnValidationError(object sender, ValidationErrorEventArgs e)
        {
            if (e.Action == ValidationErrorEventAction.Added)
                _noOfErrorsOnScreen++;
            else
                _noOfErrorsOnScreen--;

            Esegui.IsEnabled = _noOfErrorsOnScreen > 0 ? false : true;
        }
    }
}
