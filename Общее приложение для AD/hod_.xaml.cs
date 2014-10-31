using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace Общее_приложение_для_AD
{
    /// <summary>
    /// Логика взаимодействия для hod_.xaml
    /// </summary>
    public partial class hod_ : Window
    {
        DispatcherTimer DT;
        int znach = 0;
        int vse = 1;
        public hod_()
        {
            InitializeComponent();
            this.DT = new DispatcherTimer();
            this.DT.Tick += DT_Tick;
            this.DT.Interval = new TimeSpan(5);
           // this.DT.Start();
        }

        void DT_Tick(object sender, EventArgs e)
        {
            ZagruzkaSPIS.Content = String.Format("Идет загрузка списка. {0} из {1} ", znach, vse);
            Progress_ZagruzkaSPIS.Maximum = vse;
            Progress_ZagruzkaSPIS.Value = znach;
            if(vse == znach){
                this.DT.Stop();
            }
        }
        public void Setup_param1(int znach, int vse){
            this.znach = znach;
            this.vse = vse;
  
        }
        public void Setup_param(int znach, int vse)
        {
            Dispatcher.BeginInvoke(new ThreadStart(delegate {
                ZagruzkaSPIS.Content = String.Format("Идет загрузка списка. {0} из {1} ", znach, vse);
                Progress_ZagruzkaSPIS.Maximum = vse;
                Progress_ZagruzkaSPIS.Value = znach;
            }));
            BackgroundWorker BW;

        }
    }
}
