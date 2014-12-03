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
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;
using Microsoft.Win32;
using System.Security.Principal;
using System.Windows.Threading;
using System.Threading;

namespace Общее_приложение_для_AD
{


    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public csv_интерпритация.Excel_work EX;
        System.Timers.Timer Tomer_for_hod;
        DispatcherTimer DT_hod;
        hod_ HOD;
        Thread T_hod, zagruzka_spiska_perv;

        to_doc.user_card asdf;
        public String file_name;
        int CURRENT = 0, MAXIMUM = 0;
        void init_DT()
        {
            this.DT_hod = new DispatcherTimer(/*DispatcherPriority.Background*/);
            this.DT_hod.Interval = new TimeSpan(1);
            this.DT_hod.Tick += DT_hod_Tick;
        }
        public MainWindow()
        {
            InitializeComponent();
            this.button3.IsEnabled = false;


            //   this.DT_hod.IsEnabled = true;
            // this.Tomer_for_hod = new Timer( new TimerCallback(TTC),);

            this.Tomer_for_hod = new System.Timers.Timer(3);
            this.Tomer_for_hod.Elapsed += Tomer_for_hod_Elapsed;



            this.asdf = new to_doc.user_card();
            if (this.asdf.ctx == null)
            {
                MessageBox.Show("Нет доступа к домену!!!");
                Close();

            }
            else
            {
               
            }
        }

        void Tomer_for_hod_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (this.HOD != null)
            {
                this.HOD.Setup_param(this.CURRENT, this.MAXIMUM);
            }
            else
            {
                // this.HOD = new hod_();
                // this.HOD.Show();
            }
        }

        private void TTC(object state)
        {
            if (this.HOD != null)
            {
                this.HOD.Setup_param(this.CURRENT, this.MAXIMUM);
            }
            else
            {
                this.HOD = new hod_();
                this.HOD.Show();
            }
        }

        void DT_hod_Tick(object sender, EventArgs e)
        {
            if (this.HOD != null)
            {
                this.HOD.Setup_param(this.CURRENT, this.MAXIMUM);
            }
            else
            {
                this.HOD = new hod_();
                this.HOD.Show();
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog OFD = new OpenFileDialog();
            // OFD.ShowDialog();
            var BL = (bool)OFD.ShowDialog();
            if (BL == true)
            {
                this.button3.IsEnabled = true;
                this.file_name = OFD.FileName;
                PUT1.Content = OFD.SafeFileName;

            }
            else
            {
                this.file_name = "";
            }

        }
        void Init_hod_window()
        {
            this.HOD = new hod_();
            HOD.Show();
        }
        Thread TH, load_users;

       
        void INIT_HOD()
        {
            this.TH = new Thread(Init_hod_window);
            this.TH.SetApartmentState(ApartmentState.STA);
            this.TH.Start();



        }
        List<to_doc.Users> List_USERS_in_gruop;
      
        /*пользователи*/
        private void namess_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
       
        object GetParam(PrincipalContext ctx, string StrokaPodkluch, string poluchaemoe)
        {
            UserPrincipal foundUser1 = UserPrincipal.FindByIdentity(ctx, StrokaPodkluch);
            string temp = null;
            var e11 = (DirectoryEntry)foundUser1.GetUnderlyingObject();
            if ((e11.Properties[poluchaemoe]).Value != null)
                temp = (e11.Properties[poluchaemoe]).Value.ToString();
            else
                temp = "";
            return temp;
        }

        /*
         * задумка
         * хранить две строки - CN человека
         * название отдела (врятли.. но)
         * начальника CN
         */

        string test_obs(DirectoryEntry e1, string dan)
        {
            if (dan.CompareTo("objectsid") == 0)
            {
                var sidInBytes = (byte[])(e1.Properties[dan]).Value;
                var sid = new SecurityIdentifier(sidInBytes, 0);
                // This gives you what you want
                return sid.ToString();
            }
            else
                if ((e1.Properties[dan]).Value != null)
                    return (e1.Properties[dan]).Value.ToString();
                else
                    return "";
        }
        void funct2(string firstname, string lastname)
        {
            string DomainPath = to_doc.user_card.GetDomainFullName(Environment.UserDomainName);
            DirectoryEntry searchRoot = new DirectoryEntry("LDAP://" + DomainPath);
            DirectorySearcher d = new DirectorySearcher(searchRoot);
            d.Filter = string.Format("(&(objectCategory=person)(objectClass=user)(givenname={0})(sn={1}))", firstname, lastname);
            d.PropertiesToLoad.Add("name");
            d.PropertiesToLoad.Add("cn");
            d.PropertiesToLoad.Add("sn");
            d.PropertiesToLoad.Add("manager");
            var result = d.FindAll();


        }
        Thread TH1;
        DispatcherTimer DT;
        private void button3_Click(object sender, RoutedEventArgs e)
        {
            this.DT = new DispatcherTimer();
            this.DT.Interval = new TimeSpan(0, 0, 0, 0, 2);
            this.DT.Tick += DT_Tick;
            this.DT.Start();

            if (this.file_name != null)
            {
                //button3.IsEnabled = false;
                this.TH1 = new Thread(to_excel_thread);
                this.TH1.SetApartmentState(ApartmentState.STA);
                this.TH1.Start(null);
                // TH.Join();
                //   button3.IsEnabled = true;
            }
        }

        private void DT_Tick(object sender, EventArgs e)
        {
            // this.HOD = new hod_();
            try
            {
                this.button3.IsEnabled = false;
                progress__.Maximum = this.EX.BD_c.list_of_users.Count;
                progress__.Value = this.EX.vot;
                if (progress__.Maximum == progress__.Value)
                {
                    this.button3.IsEnabled = true;
                }
            }
            catch (Exception e1)
            {

            }

            //throw new NotImplementedException();
        }

        private void to_excel_thread(object obj)
        {
            this.EX = new csv_интерпритация.Excel_work();
            this.EX.csv_to_DB(this.file_name);
            var test = this.EX.BD_c;
            this.EX.ff_osn();
            this.button3.Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                {
                    this.button3.IsEnabled = true;
                }
                    )
                );
        }

        private void image1_ImageFailed(object sender, ExceptionRoutedEventArgs e)
        {

        }
        //Выход
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {               
                this.TH1.Abort();
            }
            catch (Exception er)
            {
            }
            Application.Current.MainWindow.Close();
        }

        private void STOP__Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.TH1.Abort(); 
                this.EX.vot = 0;
                this.STOP_.IsEnabled = false;
                this.button3.IsEnabled = false;
            }
            catch (Exception er)
            {
            }
        }
      
    }
}
