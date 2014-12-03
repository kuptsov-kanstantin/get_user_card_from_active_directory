using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.IO;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;
namespace csv_интерпритация
{
    public class Excel_work
    {
        public BD_class BD_c;
        private class reading_users_from_csv
        {

        }
        /*перевод из csv в базу данных*/
        public void csv_to_DB(string put)
        {
            try
            {
                this.BD_c = new BD_class();
                var VAL = File.ReadAllBytes(put);
                String values = File.ReadAllText(put, System.Text.Encoding.Default);
                var ad = values.Split('\r', '\n');
                Array.Sort(ad);
                int fd = 0;
                for (int i = 0; i < ad.Length; i++)
                {
                    if (String.Compare(ad[i], "") != 0)
                    {
                        fd = i;
                        break;
                    }
                }
                var new_str = new String[ad.Length - fd];
                for (int i = 0; i < new_str.Length; i++)
                {
                    new_str[i] = ad[fd + i];
                }
                var saf = new String[new_str.Length, 3];
                for (int i = 0; i < new_str.Length; i++)
                {
                    if (i == 2035)
                    {
                    }
                    var hhahhi = new_str[i].Split(';');
                    for (int j = 0; j < 3; j++)
                    {
                        saf[i, j] = hhahhi[j];
                    }
                    //фамилия  //день    //время
                    if (i - 1 >= 0)
                    {
                        //сравнение с фамилией.
                        if (String.Compare(saf[i - 1, 0], saf[i, 0]) == 0)
                        {
                            //сравнение с датой.
                            if (String.Compare(saf[i - 1, 1], saf[i, 1]) == 0)
                            {
                                this.BD_c.Add(saf[i, 2]);
                            }
                            else
                            {
                                this.BD_c.Add(saf[i, 1], saf[i, 2]);
                            }
                        }
                        else
                        {
                            this.BD_c.update();
                            this.BD_c.Add(saf[i, 0], saf[i, 1], saf[i, 2]);
                        }
                    }
                    else
                    {
                        this.BD_c.Add(saf[i, 0], saf[i, 1], saf[i, 2]);
                    }
                }
                this.BD_c.update();  
            }
            catch (Exception e1)
            {
                //      System.Windows.MessageBox.Show(e1.ToString(), "Ошибка");
            }


        }

        public Excel.Workbook wb;
      public  Excel.Application xlApp;
        private string for_initials(String NAME, String FAM, String OTH)
        {
            if (NAME != null)
            {
                if (FAM != null)
                {
                    if (OTH != null)
                    {
                        return NAME[0] + ". " + OTH[0] + ". " + FAM;
                    }
                    else
                    {
                        return NAME[0] + ". " + FAM;

                    }
                }
            }
            return null;
        }
        private string for_initials_n(String NAME, String FAM, String OTH)
        {
            if (NAME != null)
            {
                if (FAM != null)
                {
                    if (OTH != null)
                    {
                        return FAM+" "+ NAME[0] + ". " + OTH[0] + ". ";
                    }
                    else
                    {
                        return FAM + " "+ NAME[0] + ". " ;

                    }
                }
            }
            return null;
        }
        /*формирование страницы*/
        private Excel.Worksheet ListSheets(/*int id, */Excel.Worksheet WS_exc, fam_class DATA, to_doc.NAME_id USER_info)
        {
            int index = 0;
            //    String Famil = "", ima = "", otchestvo = "";
            String DAY_to_BEGIN = "06:00:00", FIO;
            if (USER_info != null)
            {
                FIO = for_initials(USER_info.name, USER_info.family, USER_info.oth);
                DAY_to_BEGIN = "6:00:00";
            }
            else
            {
                FIO = DATA.familia;
            }
            WS_exc.get_Range(String.Format("A1"), Type.Missing).EntireColumn.ColumnWidth = 15;
            WS_exc.get_Range(String.Format("B1"), Type.Missing).EntireColumn.ColumnWidth = 15;
            WS_exc.get_Range(String.Format("C1"), Type.Missing).EntireColumn.ColumnWidth = 15;
            WS_exc.get_Range(String.Format("D1"), Type.Missing).EntireColumn.ColumnWidth = 15;
            WS_exc.get_Range(String.Format("A5"), Type.Missing).EntireRow.RowHeight = 30;


            WS_exc.get_Range("A1", Type.Missing).Font.Bold = true;
            WS_exc.get_Range("A1", Type.Missing).Value2 = FIO;
            WS_exc.get_Range("C1", Type.Missing).Value2 = "Дата";
            WS_exc.get_Range("C2", Type.Missing).Value2 = "Норма времени";
            WS_exc.get_Range("D2", Type.Missing).Value2 = DAY_to_BEGIN;
            WS_exc.get_Range("C3", Type.Missing).Value2 = "Обед";
            WS_exc.get_Range("D3", Type.Missing).Value2 = String.Format("0:45:00");
            WS_exc.get_Range("A5", Type.Missing).Value2 = "Дата";
            WS_exc.get_Range("A5", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            WS_exc.get_Range("A5", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;
            WS_exc.get_Range("A5", Type.Missing).WrapText = true;
            WS_exc.get_Range("B5", Type.Missing).Value2 = "Начало рабочего дня";
            WS_exc.get_Range("B5", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            WS_exc.get_Range("B5", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;
            WS_exc.get_Range("B5", Type.Missing).WrapText = true;
            WS_exc.get_Range("C5", Type.Missing).Value2 = "Конец рабочего дня";
            WS_exc.get_Range("C5", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            WS_exc.get_Range("C5", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;
            WS_exc.get_Range("C5", Type.Missing).WrapText = true;
            WS_exc.get_Range("D5", Type.Missing).Value2 = "Отработано часов";
            WS_exc.get_Range("D5", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            WS_exc.get_Range("D5", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;
            WS_exc.get_Range("D5", Type.Missing).WrapText = true;
            double per_dt = 0;
            var count_bd = DATA.data_lists.Count;
            for (int data = 0; data < count_bd; data++)
            {
                var perem = DATA.data_lists[data];
                if (perem.data != null)
                {
                    //DATA_otch = perem.data;
                    var data_temp = Convert.ToDateTime(perem.data);

                    WS_exc.get_Range(String.Format("A{0}", data + 5), Type.Missing).Value2 = perem.data;
                    WS_exc.get_Range(String.Format("B{0}", data + 5), Type.Missing).Value2 = Convert.ToDateTime(perem.Times[0].time).TimeOfDay.ToString();//начало дня
                    WS_exc.get_Range(String.Format("C{0}", data + 5), Type.Missing).Value2 = Convert.ToDateTime(perem.Times[perem.Times.Count - 1].time).TimeOfDay.ToString();//конец
                    WS_exc.get_Range(String.Format("D{0}", data + 5), Type.Missing).FormulaLocal = String.Format(@"=ЕСЛИ(C{0}-B{0}-$D$3<=$D$3;$D$2;C{0}-B{0}-$D$3)", data + 5);
                    WS_exc.get_Range(String.Format("D{0}", data + 5), Type.Missing).NumberFormat = "HH:MM:SS";
                }
            }
            var hh1 = '"';
            WS_exc.get_Range("D1", Type.Missing).FormulaLocal = String.Format("=ТЕКСТ(A6;{0}ММММ ГГГГ{0})", hh1);
            WS_exc.get_Range("D1", Type.Missing).NumberFormat = "MM YY";
            WS_exc.get_Range("D1", Type.Missing).Font.Bold = true;

            WS_exc.get_Range(String.Format("C{0}", count_bd + 5), Type.Missing).Value2 = "Всего";
            WS_exc.get_Range(String.Format("D{0}", count_bd + 5), Type.Missing).FormulaLocal = String.Format("=СУММ(D6:D{0})", count_bd + 4);
            WS_exc.get_Range(String.Format("D{0}", count_bd + 5), Type.Missing).NumberFormat = "[HH]:MM";
            WS_exc.get_Range(String.Format("C{0}", count_bd + 5 + 1 + 3), Type.Missing).Value2 = "Итого";
            WS_exc.get_Range(String.Format("D{0}", count_bd + 5 + 1 + 3), Type.Missing).FormulaLocal = String.Format("=ОКРУГЛ(СУММ(D{0}:D{1})*24;0)", count_bd + 5, count_bd + 5 + 3);
            /// подписи
            WS_exc.get_Range(String.Format("B{0}", count_bd + 5 + 6), Type.Missing).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            WS_exc.get_Range(String.Format("B{0}", count_bd + 5 + 6), Type.Missing).Font.Bold = true;
            WS_exc.get_Range(String.Format("C{0}", count_bd + 5 + 6), Type.Missing).FormulaLocal = String.Format("=ЕСЛИ(A1<>" + hh1 + hh1 + ";A1;" + hh1 + "Сотрудник" + hh1 + ")");
            WS_exc.get_Range(String.Format("C{0}", count_bd + 5 + 6), Type.Missing).Font.Bold = true;


            WS_exc.get_Range(String.Format("B{0}", count_bd + 5 + 8), Type.Missing).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            WS_exc.get_Range(String.Format("B{0}", count_bd + 5 + 8), Type.Missing).Font.Bold = true;
            if (USER_info != null)
            {// 21 или 23 или 4 
                WS_exc.get_Range(String.Format("C{0}", count_bd + 5 + 8), Type.Missing).FormulaLocal = String.Format("{0}", for_initials(USER_info.nach_of_depart_name, USER_info.nach_of_depart_family, USER_info.nach_of_depart_oth));
                WS_exc.get_Range(String.Format("C{0}", count_bd + 5 + 8), Type.Missing).Font.Bold = true;
            }
            else
            {
                WS_exc.get_Range(String.Format("C{0}", count_bd + 5 + 8), Type.Missing).FormulaLocal = String.Format("{0}", "Начальник отдела");
                WS_exc.get_Range(String.Format("C{0}", count_bd + 5 + 8), Type.Missing).Font.Bold = true;
            }
                try
                {
                    WS_exc.Name = for_initials_n(USER_info.name, USER_info.family,USER_info.oth);
                }
                catch (Exception E)
                {

                }
            return WS_exc;
        }

        PrincipalContext ctx;
        GroupPrincipal grp;
        /*типа main))*/
        public int vot = 0;
        public void ff_osn()
        {
            this.xlApp = new Excel.Application("test.xsxl");
            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            wb = this.xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            for (int i = 0; i < this.BD_c.list_of_users.Count; i++)
            {
                this.vot = i;
                var USER__ = this.BD_c.list_of_users[i];
                var familia_imya = USER__.familia;
                var two  = to_doc.NAME_id.return_fam_name_otch(0, familia_imya);
                var one = to_doc.NAME_id.return_fam_name_otch(1, familia_imya);
                var rez = to_doc.user_card.get_ima_fam(one, two);
                var ws_n = (Excel.Worksheet)this.xlApp.Worksheets.Add();
                ws_n = this.ListSheets(ws_n, USER__, rez);
            }
            this.xlApp.Visible = true;//// делает видимым окно экселя...   */
        }
    }



}
