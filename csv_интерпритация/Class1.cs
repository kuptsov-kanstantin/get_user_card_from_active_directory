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
        private class reading_users_from_csv {
            
        }
        /*перевод из csv в базу данных*/
        void csv_to_DB(string put)
        {
            try
            {
                String values = File.ReadAllText(put);
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
                        //сравнивае с фамилии.
                        if (String.Compare(saf[i - 1, 0], saf[i, 0]) == 0)
                        {
                            //сравнивае с датой.
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
               //  this.update_list();

                //  var sd = new wpf_подсчет_excel.CreateExcelWorksheet();
                //   sd.CreateExcelWorksheet1();
            }
            catch (Exception e1)
            {
                //      System.Windows.MessageBox.Show(e1.ToString(), "Ошибка");
            }


        }

        Excel.Workbook wb;
        Excel.Application xlApp;

        /*формирование страницы*/
        private Excel.Worksheet ListSheets(/*int id, */Excel.Worksheet WS_exc, rab rabotnici, fam_class BD_f, string kto_podpis)
        {
            int index = 0;
            //    String Famil = "", ima = "", otchestvo = "";
            String DAY_to_BEGIN = "06:00:00", FIO;
            if (rabotnici != null)
            {
                /* Famil = rabotnici.Familia;
                 ima = rabotnici.name + ". ";
                 otchestvo = rabotnici.otchestvo + ". ";*/
                FIO = null /*Window1.for_initials(rabotnici)*/;
                DAY_to_BEGIN = rabotnici.DT_rabochi_den.hour + ":" + rabotnici.DT_rabochi_den.minute + ":" + rabotnici.DT_rabochi_den.second;
            }
            else
            {
                FIO = BD_f.familia;
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
            WS_exc.get_Range("A5", Type.Missing).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            WS_exc.get_Range("A5", Type.Missing).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            WS_exc.get_Range("A5", Type.Missing).WrapText = true;
            WS_exc.get_Range("B5", Type.Missing).Value2 = "Начало рабочего дня";
            WS_exc.get_Range("B5", Type.Missing).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            WS_exc.get_Range("B5", Type.Missing).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            WS_exc.get_Range("B5", Type.Missing).WrapText = true;
            WS_exc.get_Range("C5", Type.Missing).Value2 = "Конец рабочего дня";
            WS_exc.get_Range("C5", Type.Missing).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            WS_exc.get_Range("C5", Type.Missing).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            WS_exc.get_Range("C5", Type.Missing).WrapText = true;
            WS_exc.get_Range("D5", Type.Missing).Value2 = "Отработано часов";
            WS_exc.get_Range("D5", Type.Missing).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            WS_exc.get_Range("D5", Type.Missing).VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            WS_exc.get_Range("D5", Type.Missing).WrapText = true;
            double per_dt = 0;
            var count_bd = BD_f.D_T_l.Count;
            for (int data = 0; data < count_bd; data++)
            {
                var perem = BD_f.D_T_l[data];
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
            if (rabotnici != null)
            {// 21 или 23 или 4 

                WS_exc.get_Range(String.Format("C{0}", count_bd + 5 + 8), Type.Missing).FormulaLocal = String.Format("{0}", kto_podpis);
                WS_exc.get_Range(String.Format("C{0}", count_bd + 5 + 8), Type.Missing).Font.Bold = true;
            }
            else
            {
                WS_exc.get_Range(String.Format("C{0}", count_bd + 5 + 8), Type.Missing).FormulaLocal = String.Format("{0}", "Начальник отдела");
                WS_exc.get_Range(String.Format("C{0}", count_bd + 5 + 8), Type.Missing).Font.Bold = true;
            }
            //=ЕСЛИ(A1<>"";A1;"Сотрудник")

            if (rabotnici != null)
            {
               // WS_exc.Name = Window1.for_initials_n(rabotnici);//<<< Переименование листа!!!!! <<<<<<< 
            }
            else
            {
                WS_exc.Name = BD_f.familia;
            }

            return WS_exc;
        }
       
        PrincipalContext ctx;
        GroupPrincipal grp;
        /*типа main))*/
        public void ff_osn()
        {
        



            /*



            this.xlApp = new Excel.Application("test.xsxl");
            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            wb = this.xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            int zn = 0;
            var TRANSLIT = BD_c.D_T;
            for (int id = 0; id < TRANSLIT.Count; id++)
            {
                var TRANSLIT_name = TRANSLIT[id];
                if (rsa != null)
                {
                    bool asnaeb = false;
                    for (int yy = 0; yy < rsa.Count; yy++)//поиск сотрудника в БД
                    {
                        var ima_bd = rsa[yy];
                        if (String.Compare(TRANSLIT_name.familia, ima_bd.TRANSLIT) == 0)///сопоставление с базой
                        {
                            this.fiajsdkfjh = id;
                            string podpis = null;
                            var ws_n = (Excel.Worksheet)this.xlApp.Worksheets.Add();
                            for (int yt = 0; yt < rsa.Count; yt++)//поиск начальника
                            {
                                if (yy != yt)
                                {
                                    var rsa_nach = rsa[yt];
                                    if (ima_bd.otdel == rsa_nach.otdel)
                                    {
                                        if (rsa_nach.Naxalnik_otd_da_net == true)
                                        {
                                            podpis = Window1.for_initials(rsa_nach);
                                            asnaeb = true;
                                            break;
                                        }
                                    }
                                }
                                else
                                {
                                    var rsa_nach = rsa[yt];
                                    if (ima_bd.otdel == rsa_nach.otdel)
                                    {
                                        if (rsa_nach.Naxalnik_otd_da_net == true)
                                        {
                                            podpis = "П. Котэк";
                                            asnaeb = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (asnaeb == false)
                            {
                                podpis = "Начальник отдела >>  " + rab.obj_((List_obj)savii.open_class("otdeli.xml", typeof(List_obj)), ima_bd.otdel);
                            }
                            asnaeb = true;

                            ws_n = this.ListSheets(ws_n, ima_bd, TRANSLIT_name, podpis);

                        }
                    }
                    if (asnaeb == false)
                    {
                        this.fiajsdkfjh = id;
                        var ws_n = (Excel.Worksheet)this.xlApp.Worksheets.Add();
                        ws_n = this.ListSheets(ws_n, null, TRANSLIT_name, "Начальник отдела");
                    }
                }
                else
                {
                    var aaa = this.BD_c.D_T[id];
                    this.fiajsdkfjh = id;
                    var ws_n = (Excel.Worksheet)this.xlApp.Worksheets.Add();
                    ws_n = this.ListSheets(ws_n, null, aaa, "Начальник отдела");
                }
            }
            this.xlApp.Visible = true;//// делает видимым окно экселя...      
            // save();
            this.timer.Stop();*/
        }
    }



}
