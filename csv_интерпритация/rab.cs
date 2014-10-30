using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace csv_интерпритация
{

    using List_obj = System.Collections.ObjectModel.ObservableCollection<Object_i>;
    using rab_list = System.Collections.ObjectModel.ObservableCollection<rab>;

    static public class savii
    {
        public static void save_calss(String file, Object myObject)
        {
            try
            {
                var mySerializer = new XmlSerializer(myObject.GetType());
                var myWriter = new StreamWriter(file);
                mySerializer.Serialize(myWriter, myObject);
                myWriter.Close();
            }
            catch (Exception er)
            {
                //System.Windows.MessageBox.Show(er.ToString(), "Ошибка");
            }
        }
        //"rabotniki.xml"
        public static Object open_class(String file, Type types)
        {
            try
            {
                var myWriter = new StreamReader(file);
                var mySerializer = new XmlSerializer(types);
                var myObject = (mySerializer.Deserialize(myWriter));
                myWriter.Close();
                return myObject;
            }
            catch (FileNotFoundException er)
            {
                //System.Windows.MessageBox.Show(er.ToString(), "Ошибка");
            }
            catch (Exception er)
            {
                //System.Windows.MessageBox.Show(er.ToString(), "Ошибка");
            }  
            return null;
        }
    }
    public class time_class
    {
        public String time;
        public time_class(String time_)
        {
            this.time = time_;
        }
    }
    //дата и время. Дата как основа. Время добавляемое 
    public class Data_time_class
    {
        public String data;
        public List<time_class> Times;
        public Data_time_class()
        {
            this.Times = new List<time_class>();
        }
        //добавление новой даты со временем
        public Data_time_class(String data_, String time_)
        {
            this.data = data_;
            if (this.Times == null)
            {
                this.Times = new List<time_class>();
            }
            this.Times.Add(new time_class(time_));
        }
        //добавление времени к текущей дате
        public Data_time_class(String time_)
        {
            this.Times.Add(new time_class(time_));
        }
        public void Add_Data_time(String data_, String time_)
        {
            this.data = data_;
            if (this.Times == null)
            {
                this.Times = new List<time_class>();
            }
            this.Times.Add(new time_class(time_));
        }
        public void Add_Data_time(String time_)
        {
            this.Times.Add(new time_class(time_));
        }
    }

    // фамилия 
    public class fam_class
    {
        public String familia;
        private Data_time_class D_T;
        public List<Data_time_class> data_lists;
        public fam_class(){   }
        //новый человек. сохраняет фамилию и инициализирует время и дату.
        public fam_class(String familia_)
        {
            this.familia = familia_;
            this.D_T = new Data_time_class();
            this.data_lists = new List<Data_time_class>();
        }
        //добавление времени к текущему человеку в текущей дате.
        public void fam_cl_add(String time_)
        {
            this.D_T.Add_Data_time(time_);
        }
        //добавление новой даты вместе со временем.
        public void fam_cl_add(String data_, String time_)
        {
            this.update();
            this.D_T = new Data_time_class();
            this.D_T.Add_Data_time(data_, time_);
        }
        public void update()
        {
            this.data_lists.Add(this.D_T);
        }
        public fam_class(String familia_, String data_, String time_)
        {

        }
    }
    public class BD_class
    {
        public List<fam_class> list_of_users;
        private fam_class asd;
        public BD_class()
        {
            list_of_users = new List<fam_class>();
            asd = new fam_class();
        }
        // нового человека с фамилией... добавление к нему даты и время
        public void Add(String familia_, String data_, String time_)
        {
            this.asd = new fam_class(familia_);
            this.asd.fam_cl_add(data_, time_);
        }
        // добавление к нему даты и время
        public void Add(String data_, String time_)
        {
            this.asd.fam_cl_add(data_, time_);
        }
        //добавление времени.
        public void Add(String time_)
        {
            this.asd.fam_cl_add(time_);
        }
        //обновление
        public void update()
        {
            this.asd.update();
            this.list_of_users.Add(this.asd);
        }
    }

    public class organizacia
    {
        public lists list_dolzn_otdels;
        public class lists
        {
            public List<String> dolzn;
            public List<String> otdels;
            public lists()
            {

                this.dolzn = new List<String>();
                this.otdels = new List<String>();
            }
        }
        public organizacia()
        {
            this.list_dolzn_otdels = new lists();

        }
        public void save(String from_file)
        {
            savii.save_calss(from_file, this.list_dolzn_otdels);
        }
        public void open(String from_file)
        {
            this.list_dolzn_otdels = (lists)savii.open_class(from_file, typeof(rab_list));
        }
    }




    //public class rab_list
    //{
    //    public List<rab> LS_rab;
    //    public rab_list()
    //    {
    //        this.LS_rab = new List<rab>();
    //    }
    //}
    public class rab
    {
        public String
            Familia,
            name,
            otchestvo,
            TRANSLIT;
        public bool Naxalnik_otd_da_net = false;
          // kto_podpis;
       public int dolznost = 0, otdel = 0;
       public its_time_to_begin DT_rabochi_den;
       public rab()
       {

       }
       public rab(String Familia, String name, String otchestvo, String TRANSLIT,  its_time_to_begin DT_rabochi_den, int dolzn, int otdel, bool Naxalnik_otd_da_net)
       {
           this.Familia = Familia;
           this.name = name;
           this.otchestvo = otchestvo;
           this.TRANSLIT = TRANSLIT;           
           this.DT_rabochi_den = DT_rabochi_den;
           this.dolznost = dolzn;
           this.otdel = otdel;
           this.Naxalnik_otd_da_net = Naxalnik_otd_da_net;
       }
       public static string obj_(List_obj tes, int i)
       {
           if (i > -1)
           {
               return tes[i].OBJ;
           }
           return null;
       }
    }
    public class DATATest_otd
    {
        public string ID { get; set; }
        public string OBJ { get; set; }
    }
    public class Test
    {
        public string Data_in_month { get; set; }
        public string Begin_day { get; set; }
        public string End_of_day { get; set; }
        public string Hours_worked { get; set; }
    }
    public class Test1
    {

        public string Data_name { get; set; }
        public string Data_familia { get; set; }
        public string Data_otchestvo { get; set; }
        public string Data_FIO_transl { get; set; }
        public string Data_dolj { get; set; }
        public string Data_otd { get; set; }
        public string Data_norm { get; set; }
        public string Naxalnik_otd_da_net { get; set; }


    //    public int ID__ { get; set; }
        public Test1(rab LS_rab)
        {
            this.Data_name = LS_rab.name;
            this.Data_familia = LS_rab.Familia;
            this.Data_otchestvo = LS_rab.otchestvo;
            this.Data_FIO_transl = LS_rab.TRANSLIT;
            this.Data_dolj = rab.obj_((List_obj)savii.open_class("dolznost.xml", typeof(List_obj)), LS_rab.dolznost);
            this.Data_otd = rab.obj_((List_obj)savii.open_class("otdeli.xml", typeof(List_obj)), LS_rab.otdel); 
            this.Data_norm = String.Format("{0}:{1}:{2}", LS_rab.DT_rabochi_den.hour, LS_rab.DT_rabochi_den.minute, LS_rab.DT_rabochi_den.second);
          if(LS_rab.Naxalnik_otd_da_net == true){
              this.Naxalnik_otd_da_net = "Да";
          }
          else
          {
              this.Naxalnik_otd_da_net = "Нет";
          }
            
        }

    }

    public class Object_i
    {

        public int ID = 0;
        public String OBJ = "";
        public Object_i(int ID, String OBJ)
        {
            this.ID = ID;
            this.OBJ = OBJ;
        }
        public Object_i()
        {

        }
    }

}
