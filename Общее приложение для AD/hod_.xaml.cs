using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Общее_приложение_для_AD
{
    /// <summary>
    /// Логика взаимодействия для hod_.xaml
    /// </summary>
    public partial class hod_ : Window
    {  
        public hod_()
        {
            InitializeComponent();
        }
        public void Setup_param(int znach, int vse){
            ZagruzkaSPIS.Content = String.Format("Идет загрузка списка. {0} из {1} ", znach, vse);
            Progress_ZagruzkaSPIS.Maximum = vse;
            Progress_ZagruzkaSPIS.Value = znach;
        }
    }
}
