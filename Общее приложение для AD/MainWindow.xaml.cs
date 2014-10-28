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

namespace Общее_приложение_для_AD
{


    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        to_doc.user_card asdf;
        public String file_name;
        public MainWindow()
        {
            InitializeComponent();
            this.asdf = new to_doc.user_card();
            if (this.asdf.ctx == null) {
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


        private void button1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog OFD = new OpenFileDialog();
           // OFD.ShowDialog();
            var BL = (bool)OFD.ShowDialog();
            if (BL == true)
            {
                this.file_name = OFD.FileName;
            }
            else {
                this.file_name = "";
            }

        }
        /*групы */
        private void comboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            namess.Items.Clear();

            var GGL = this.asdf.GetAllDep();

            var department = GGL[comboBox1.SelectedIndex];
            for (int u = 0; u < this.asdf.UsersOnList.Count; u++)
            {
                MessageBox.Show(String.Format("Идет загрузка списка. {0} из {1} ", u, this.asdf.UsersOnList.Count));
                var USER = this.asdf.UsersOnList[u];
                if (String.Compare(USER.DEPARTMENT, department) == 0)
                {
                    var user2 = this.asdf.GetUSERbySID(USER.SID);
                    namess.Items.Add(user2.FIO + " | " + user2.login);
                }     
            }
            if (namess.Items.Count > 0)
            {
                namess.SelectedIndex = 0;
            }


           /* var USERs = this.asdf.GetUserList(comboBox1.SelectedIndex);
            if (USERs != null)
            {
                for (int i = 0; i < USERs.Count; i++)
                {
                    namess.Items.Add(USERs[i].FIO + " | " + USERs[i].login);
                }
                if (namess.Items.Count > 0)
                {
                    namess.SelectedIndex = 0;
                }
            }*/
        }
        /*пользователи*/
        private void namess_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            var USERs = this.asdf.GetUserList(comboBox1.SelectedIndex);
            this.asdf.HTML_to_doc(USERs[namess.SelectedIndex].FIO, USERs[namess.SelectedIndex].login, "", "", USERs[namess.SelectedIndex].mail, "");

        }
        object GetParam(PrincipalContext ctx, string StrokaPodkluch, string poluchaemoe ) {
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






     
        private void button3_Click(object sender, RoutedEventArgs e)
        {

    //  var fdf =      GetADUsers();
    




            var ctx = new PrincipalContext(ContextType.Domain, Environment.UserDomainName);
            UserPrincipal foundUser = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, "kuptsov");/// поиск пользователя
            var e1 = (DirectoryEntry)foundUser.GetUnderlyingObject();//получение информации о человеке
            String FIO, mail, login, department, manager, FIO_n;

            FIO = test_obs(e1, "cn");
            string sid = test_obs(e1, "objectsid");
            mail = test_obs(e1, "mail");
            login = test_obs(e1, "sAMAccountName");
            department = test_obs(e1, "department");

            if ((e1.Properties["manager"]).Value != null)
            {
                manager = (e1.Properties["manager"]).Value.ToString();
                UserPrincipal foundUser1 = UserPrincipal.FindByIdentity(ctx, manager);
                var e11 = (DirectoryEntry)foundUser1.GetUnderlyingObject();
                FIO_n = test_obs(e11, "cn");
            }
            else
                manager = "";
        }

        private void image1_ImageFailed(object sender, ExceptionRoutedEventArgs e)
        {

        }
    }
}
