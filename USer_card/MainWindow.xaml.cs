using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Security.Principal;
namespace USer_card
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Thread TH, load_users;
        to_doc.user_card asdf; 
        int CURRENT = 0, MAXIMUM = 0;
        public MainWindow()
        {
            InitializeComponent();

            this.asdf = new to_doc.user_card();

            if (this.asdf.ctx == null)
            {
                MessageBox.Show("Нет доступа к домену!!!");
                Close();

            }
            else
            {
                var GGL = this.asdf.GetAllDep();
                for (int i = 0; i < GGL.Count; i++)
                {
                    comboBox1.Items.Add(GGL[i]);
                }
                if (comboBox1.Items.Count > 0)
                {
                    comboBox1.SelectedIndex = 0;
                }
            }
        }
        List<to_doc.Users> List_USERS_in_gruop;
      
        /*формирование карточки*/
        private void button2_Click(object sender, RoutedEventArgs e)
        {
            var USERs = this.List_USERS_in_gruop[namess.SelectedIndex];
            var user2 = this.asdf.GetUSERbySID(USERs.SID);
            this.asdf.HTML_to_doc(user2.FIO, user2.login, "pass", "skd", user2.mail);

        }
        private void namess_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        /*групы */
        private void comboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
          //  this.INIT_HOD();
            this.load_users = new Thread(load_users_thread);
            this.load_users.SetApartmentState(ApartmentState.STA);
            this.load_users.Start();
        }
        void load_users_thread()
        {
            int pos = 0;
            Dispatcher.BeginInvoke(new ThreadStart(delegate
            {
                namess.Items.Clear();
                pos = comboBox1.SelectedIndex;
            }));
            var GGL = this.asdf.GetAllDep();
            var department = GGL[pos];
            this.List_USERS_in_gruop = new List<to_doc.Users>();
            for (int u = 0; u < this.asdf.UsersOnList.Count; u++)
            {
                this.CURRENT = u;
                this.MAXIMUM = this.asdf.UsersOnList.Count;
                var USER = this.asdf.UsersOnList[u];

                

                if (String.Compare(USER.DEPARTMENT, department) == 0)
                {
                    this.List_USERS_in_gruop.Add(new to_doc.Users(USER));
                    var user2 = this.asdf.GetUSERbySID(USER.SID);
                    Dispatcher.BeginInvoke(
                        new ThreadStart(delegate
                        {
                            try
                            {
                                if (user2.nach_of_depart != null)
                                    namess.Items.Add(user2.FIO + " (" + user2.login + ") " + user2.nach_of_depart);
                                else
                                    namess.Items.Add(user2.FIO + " (" + user2.login + ") ");
                                if (namess.Items.Count > 0)
                                {
                                    namess.SelectedIndex = 0;
                                }
                            }
                            catch (Exception E)
                            {
                            }
                        }));
                }
            }
          
        }

        /*Выход*/
        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }




    }
}
