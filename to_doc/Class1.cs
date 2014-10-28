﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;
using System.Diagnostics;
using System.Collections;
//using Json;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Web.Script.Serialization;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Security.Principal;

namespace to_doc
{
    public static class ADProperties
    {
        public const String OBJECTCLASS = "objectClass";
        public const String CONTAINERNAME = "cn";
        public const String LASTNAME = "sn";
        public const String COUNTRYNOTATION = "c";
        public const String CITY = "l";
        public const String STATE = "st";
        public const String TITLE = "title";
        public const String POSTALCODE = "postalCode";
        public const String PHYSICALDELIVERYOFFICENAME = "physicalDeliveryOfficeName";
        public const String FIRSTNAME = "givenName";
        public const String MIDDLENAME = "initials";
        public const String DISTINGUISHEDNAME = "distinguishedName";
        public const String INSTANCETYPE = "instanceType";
        public const String WHENCREATED = "whenCreated";
        public const String WHENCHANGED = "whenChanged";
        public const String DISPLAYNAME = "displayName";
        public const String USNCREATED = "uSNCreated";
        public const String MEMBEROF = "memberOf";
        public const String USNCHANGED = "uSNChanged";
        public const String COUNTRY = "co";
        public const String DEPARTMENT = "department";
        public const String COMPANY = "company";
        public const String PROXYADDRESSES = "proxyAddresses";
        public const String STREETADDRESS = "streetAddress";
        public const String DIRECTREPORTS = "directReports";
        public const String NAME = "name";
        public const String OBJECTGUID = "objectGUID";
        public const String USERACCOUNTCONTROL = "userAccountControl";
        public const String BADPWDCOUNT = "badPwdCount";
        public const String CODEPAGE = "codePage";
        public const String COUNTRYCODE = "countryCode";
        public const String BADPASSWORDTIME = "badPasswordTime";
        public const String LASTLOGOFF = "lastLogoff";
        public const String LASTLOGON = "lastLogon";
        public const String PWDLASTSET = "pwdLastSet";
        public const String PRIMARYGROUPID = "primaryGroupID";
        public const String OBJECTSID = "objectSid";
        public const String ADMINCOUNT = "adminCount";
        public const String ACCOUNTEXPIRES = "accountExpires";
        public const String LOGONCOUNT = "logonCount";
        public const String LOGINNAME = "sAMAccountName";
        public const String SAMACCOUNTTYPE = "sAMAccountType";
        public const String SHOWINADDRESSBOOK = "showInAddressBook";
        public const String LEGACYEXCHANGEDN = "legacyExchangeDN";
        public const String USERPRINCIPALNAME = "userPrincipalName";
        public const String EXTENSION = "ipPhone";
        public const String SERVICEPRINCIPALNAME = "servicePrincipalName";
        public const String OBJECTCATEGORY = "objectCategory";
        public const String DSCOREPROPAGATIONDATA = "dSCorePropagationData";
        public const String LASTLOGONTIMESTAMP = "lastLogonTimestamp";
        public const String EMAILADDRESS = "mail";
        public const String MANAGER = "manager";
        public const String MOBILE = "mobile";
        public const String PAGER = "pager";
        public const String FAX = "facsimileTelephoneNumber";
        public const String HOMEPHONE = "homePhone";
        public const String MSEXCHUSERACCOUNTCONTROL = "msExchUserAccountControl";
        public const String MDBUSEDEFAULTS = "mDBUseDefaults";
        public const String MSEXCHMAILBOXSECURITYDESCRIPTOR = "msExchMailboxSecurityDescriptor";
        public const String HOMEMDB = "homeMDB";
        public const String MSEXCHPOLICIESINCLUDED = "msExchPoliciesIncluded";
        public const String HOMEMTA = "homeMTA";
        public const String MSEXCHRECIPIENTTYPEDETAILS = "msExchRecipientTypeDetails";
        public const String MAILNICKNAME = "mailNickname";
        public const String MSEXCHHOMESERVERNAME = "msExchHomeServerName";
        public const String MSEXCHVERSION = "msExchVersion";
        public const String MSEXCHRECIPIENTDISPLAYTYPE = "msExchRecipientDisplayType";
        public const String MSEXCHMAILBOXGUID = "msExchMailboxGuid";
        public const String NTSECURITYDESCRIPTOR = "nTSecurityDescriptor";
    }
    public class Users
    {
        public Users() { }
        public Users(Users U) {
            this.SID = U.SID;
            this.CN = U.CN;
            this.DEPARTMENT = U.DEPARTMENT;
        }
        public string CN { get; set; }
        public string SID { get; set; }
        public string DEPARTMENT { get; set; }

    }
    public class NAME_id
    {
        public string login;
        public string FIO;
        public string mail;
        public string nach_of_depart;
        public NAME_id() { }
        public NAME_id(string login, string FIO, string mail, string nach_of_depart)
        {
            this.login = login;
            this.FIO = FIO;
            this.mail = mail;
            this.nach_of_depart = nach_of_depart;
        }
        ~NAME_id() { }
    }

    /*
     СДЕЛАТЬ ДЕРЕВО
     */

    public class user_card
    {
        public user_card() { this.init(); }
        ~user_card() { }
        internal sealed class SystemCore_EnumerableDebugView
        {
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private object[] cachedCollection;
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private int count;
            [DebuggerBrowsable(DebuggerBrowsableState.Never)]
            private IEnumerable enumerable;
            public SystemCore_EnumerableDebugView(IEnumerable enumerable)
            {
                if (null == enumerable) throw new ArgumentNullException("enumerable");
                this.enumerable = enumerable;
            }

            [DebuggerBrowsable(DebuggerBrowsableState.RootHidden)]
            public object[] Items
            {
                get
                {
                    var list = new List<object>();
                    IEnumerator enumerator = this.enumerable.GetEnumerator();
                    if (enumerator != null)
                    {
                        this.count = 0;
                        while (enumerator.MoveNext())
                        {
                            list.Add(enumerator.Current);
                            this.count++;
                        }
                    }
                    this.cachedCollection = new object[this.count];
                    list.CopyTo(this.cachedCollection, 0);
                    return this.cachedCollection;
                }
            }
        }
        [DataContract]
        public class Person
        {
            [DataMember]
            public string PAS;
            [DataMember]
            public int SKD;
            public Person() { }
            public Person(string PAS, int SKD)
            {
                this.SKD = SKD;
                this.PAS = PAS;
            }
        }
        static Word.Application word;
        static Word.Document wordDoc;
        public void HTML_to_doc(string FIO, string login, string pass, string SKD, string post, string argv)
        {

           
            var filepath = File.OpenText("..\\..\\HTMLPage1.html");
            String tesvt = filepath.ReadToEnd();
            String[] tem_z = { "$FIO", "$login", "$pass", "$skd", "$mail" };

            tesvt = tesvt.Replace(tem_z[0], FIO);
            tesvt = tesvt.Replace(tem_z[1], login);
            tesvt = tesvt.Replace(tem_z[2], pass);
            tesvt = tesvt.Replace(tem_z[3], SKD);
            tesvt = tesvt.Replace(tem_z[4], post);

            File.WriteAllText("temp.html", tesvt);
            //    Object confirmconversion = System.Reflection.Missing.Value;
            //   Object readOnly = false;
            //Object saveto = "c:\\doc.pdf";
            //Object oallowsubstitution = System.Reflection.Missing.Value;

            /*wordDoc = word.Documents.Open(ref filepath, ref confirmconversion, ref readOnly, ref oMissing,
                                          ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                          ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                          ref oMissing, ref oMissing, ref oMissing, ref oMissing);*/
            var strtty = Directory.GetCurrentDirectory();
            Object oMissing = System.Reflection.Missing.Value;
            if (word == null) word = new Word.Application();
            //  if (wordDoc == null) wordDoc = new Word.Document();

            // wordDoc = word.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            word.Visible = true;
            wordDoc = word.Documents.Open(strtty + "\\temp.html");

            //  File.Delete("temp.html");
            /* object fileFormat = Word.WdSaveFormat.wdFormatPDF;
             wordDoc.SaveAs(ref saveto, ref fileFormat, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oallowsubstitution, ref oMissing,
                            ref oMissing);*/

        }
        static void button1_Click(string FIO, string login, string pass, string SKD, string post)
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 6, 2, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 0.1F;

            // oTable.Range.Font = ;
            //oTable.Title = FIO;

            oTable.Cell(1, 1).Range.Text = FIO; oTable.Cell(1, 2).Range.Text = "login: " + login + Environment.NewLine + "pass " + pass + Environment.NewLine + "СКД " + SKD;
            oTable.Cell(2, 1).Range.Text = "Почта"; oTable.Cell(2, 2).Range.Text = post;
            oTable.Cell(3, 1).Range.Text = "Личная папка"; oTable.Cell(3, 2).Range.Text = String.Format("N:\\+{0}\\{1}D:\\work\\{0}\\ = {1}Мои документы{1}", login, Environment.NewLine, '"');
            oTable.Cell(4, 1).Range.Text = "Общие ресурсы"; oTable.Cell(4, 2).Range.Text = String.Format("Сетевые папки:{0}N:\\{1}Папки сотрудников{0}J:\\PW\\{1}\\ Проекты{0}X:\\{1}сканер (Toshiba в тех. отделе){0}B:\\{1}техническая (ПО, драйвера)", Environment.NewLine, '\t');
            oTable.Cell(5, 1).Range.Text = ""; oTable.Cell(5, 2).Range.Text = String.Format("ОРГ-ТЕХНИКА:{0}Toshiba (204.85.134.18){0}МФУ Ч/Б А3/А4 в тех.отделе", Environment.NewLine, '\t');
            oTable.Cell(6, 1).Range.Text = String.Format("По всем вопросом касательно работы ПК и сети пишите по почте:{0}Кунцевич Андрей Михайлович{0}a.kuntsevich@unisneft.com", Environment.NewLine); oTable.Cell(6, 2).Range.Text = String.Format("HP DJ 500 (204.85.134.20){0}Плоттер Цветной А1 в тех.отделе", Environment.NewLine, '\t');



            //Close this form.
            //this.Close();
        }
        public PrincipalContext ctx;
        GroupPrincipal grp;
        List<String> gruops;
        List<NAME_id> users;


        /*инициализация связи с доменом*/
        public void init()
        {
            try
            {
                this.ctx = new PrincipalContext(ContextType.Domain, Environment.UserDomainName);
            }
            catch (Exception e)
            {
                this.ctx = null;
            }
        }

        /**/
        string test_obs(DirectoryEntry e1, string dan)
        {

            if ((e1.Properties[dan]).Value != null)
                return (e1.Properties[dan]).Value.ToString();
            else
                return "";
        }

        /*Получение списка пользователей из группы*/
        public List<NAME_id> GetUserList(int grups_id)
        {
            if (this.gruops == null) return null;
            this.grp = GroupPrincipal.FindByIdentity(this.ctx, IdentityType.SamAccountName, this.gruops[grups_id]);
            this.users = new List<NAME_id>();
            if (grp != null)
            {
                int gg = 0;
                var tt = grp.GetMembers(true);
                this.grp.GetUnderlyingObject();
                var f = grp.Members;
                foreach (Principal p in tt)
                {
                    var e = (DirectoryEntry)p.GetUnderlyingObject();                 
                    this.users.Add(
                        new NAME_id(
                            test_obs(e, "sAMAccountName"),
                            test_obs(e, "cn"),
                            test_obs(e, "mail"), 
                            test_obs(e, "manager")
                            )
                            );
                }
            }
            return this.users;
        }
        private string GetDomainFullName(string friendlyName)
        {
            DirectoryContext context = new DirectoryContext(DirectoryContextType.Domain, friendlyName);
            Domain domain = Domain.GetDomain(context);
            return domain.Name;
        }



        /*НАВЕРНО ТУТ БУДУ ИСКАТЬ ЧЕЛА ПО SID*/
        /*Получение всех пользователей*/

        public NAME_id GetUSERbySID(string SID)
        {
            var ctx = new PrincipalContext(ContextType.Domain, Environment.UserDomainName);
            UserPrincipal foundUser = UserPrincipal.FindByIdentity(ctx, IdentityType.Sid, SID);/// поиск пользователя
            var e1 = (DirectoryEntry)foundUser.GetUnderlyingObject();//получение информации о человеке
            String FIO, mail, login, department, manager, FIO_n;


       /*     var USER = new NAME_id(test_obs(e1, "sAMAccountName"), test_obs(e1, "cn"), test_obs(e1, "mail"), 
                
                );*/




            /*
            FIO = test_obs(e1, "cn");
            string sid = test_obs(e1, "objectsid");
            mail = test_obs(e1, "mail");
            login = test_obs(e1, "sAMAccountName");
            department = test_obs(e1, "department");
            */
            if ((e1.Properties["manager"]).Value != null)
            {


                manager = (e1.Properties["manager"]).Value.ToString();
                UserPrincipal foundUser1 = UserPrincipal.FindByIdentity(ctx, manager);
                var e11 = (DirectoryEntry)foundUser1.GetUnderlyingObject();
                FIO_n = test_obs(e11, "cn");

                return new NAME_id(test_obs(e1, "sAMAccountName"), test_obs(e1, "cn"), test_obs(e1, "mail"), FIO_n);

            }
            else
            {
                return new NAME_id(test_obs(e1, "sAMAccountName"), test_obs(e1, "cn"), test_obs(e1, "mail"), "Директор");

            }

            

        }





        public List<Users> GetALLUsers()
        {
            try
            {
                List<Users> lstADUsers = new List<Users>();
                string DomainPath = GetDomainFullName(Environment.UserDomainName);
                DirectoryEntry searchRoot = new DirectoryEntry("LDAP://" + DomainPath);
                DirectorySearcher search = new DirectorySearcher(searchRoot);
                search.Filter = "(&(objectClass=user)(objectCategory=person))";
                search.PropertiesToLoad.Add("cn");
                search.PropertiesToLoad.Add("objectsid");
                //  search.PropertiesToLoad.Add("usergroup");
                search.PropertiesToLoad.Add("department");//first name
                SearchResult result;
                SearchResultCollection resultCol = search.FindAll();
                if (resultCol != null)
                {
                    for (int counter = 0; counter < resultCol.Count; counter++)
                    {
                        string UserNameEmailString = string.Empty;
                        result = resultCol[counter];
                        if (result != null)
                        {
                            Users US = new Users();
                            if (result.Properties.Contains("cn") == true)
                            {
                                US.CN = (String)result.Properties["cn"][0];
                            }
                            else
                            {
                                US.CN = "";
                            }
                            var sidInBytes = (byte[])(result.Properties["objectsid"][0]);
                            if (result.Properties.Contains("objectsid") == true)
                            {
                                US.SID = new SecurityIdentifier(sidInBytes, 0).ToString();
                            }
                            else
                            {
                                US.SID = "";
                            }
                            // This gives you what you want
                            if (result.Properties.Contains("department") == true)
                            {
                                US.DEPARTMENT = (String)result.Properties["department"][0];
                            }
                            else
                            {
                                US.DEPARTMENT = "";
                            }
                            lstADUsers.Add(US);
                        }
                    }
                }
                return lstADUsers;
            }
            catch (Exception ex)
            {

            }
            return null;
        }

       public  List<Users> UsersOnList;
        List<String> dep;
        /*получение списка отделов... теперь!*/
        public List<String> GetAllDep()
        {
            this.dep = new List<string>();
            this.UsersOnList = this.GetALLUsers();

            var temp = new List<String>();
            for (int i = 0; i < this.UsersOnList.Count; i++)
            {
                temp.Add(this.UsersOnList[i].DEPARTMENT); 
            
            }
       
            temp.Sort();
            for (int i = 0; i < temp.Count - 1; i++)
            {
                if (String.Compare(temp[i], temp[i + 1]) != 0)
                {
                    this.dep.Add(temp[i]);
                }
            }
            this.dep.Add(temp[temp.Count - 1]);
            return this.dep;
        }

        /*групп*/
        public List<String> GetGruopList()
        {
            List<String> hhf = new List<String>();
            var t = Environment.UserDomainName;
            using (var searcher = new PrincipalSearcher(new GroupPrincipal(new PrincipalContext(ContextType.Domain, t))))
            {
                foreach (var result in searcher.FindAll())
                {

                    hhf.Add(result.Name);
                }
            }
            this.gruops = hhf;
            return hhf;
        }
        void GetUserInfo(int gruop, int nomer)
        {
            this.grp = GroupPrincipal.FindByIdentity(this.ctx, IdentityType.SamAccountName, "employees");
            int yy = 0;
            var tt = this.grp.GetMembers(true);
            foreach (Principal p in tt)
            {
                if (yy == Convert.ToInt32(nomer))
                {
                    var e = (DirectoryEntry)p.GetUnderlyingObject();
                    String FIO = (e.Properties["cn"]).Value.ToString();
                    String mail = (e.Properties["mail"]).Value.ToString();
                    String desc;
                    if ((e.Properties["description"]) != null)
                    {
                        desc = (e.Properties["description"]).Value.ToString();
                    }
                    else
                    {
                        desc = " ";
                    }
                    String login = (e.Properties["sAMAccountName"]).Value.ToString();
                    if (desc != null)
                    {
                        Person p2 = new JavaScriptSerializer().Deserialize<Person>(desc.ToString());
                        Console.Write("{0} {1} {2}", desc.ToString(), p2.PAS, p2.SKD);
                        HTML_to_doc(FIO, login, p2.PAS, p2.SKD.ToString(), mail, "");
                        // button1_Click(FIO, login, p2.PAS, p2.SKD.ToString(), mail);

                    }
                }
                yy++;
            }
        }

        public void Main1(string[] args)
        {
            string groupName = "employees";
            string domainName = "adcontrol";

            //  PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domainName, "kuptsov", "BK8dzztD");

            //  UserPrincipal grp1 = new UserPrincipal(ctx);// Для создания пользователя
            if (grp != null)
            {
                int gg = 0;
                var tt = grp.GetMembers(true);
                grp.GetUnderlyingObject();
                var f = grp.Members;
                foreach (Principal p in tt)
                {
                    Console.WriteLine("{0} {1} {2} ", gg++, p.SamAccountName, p.DisplayName);
                }
                Console.WriteLine("-1 - всех{0} Ну или номер из списка", Environment.NewLine);
                string nomer = Console.ReadLine();
                int yy = 0;
                foreach (Principal p in tt)
                {
                    if (Convert.ToInt32(nomer) == -1)
                    {
                        var e = (DirectoryEntry)p.GetUnderlyingObject();
                        String FIO = (e.Properties["cn"]).Value.ToString();
                        String mail = (e.Properties["mail"]).Value.ToString();
                        String desc;
                        if ((e.Properties["description"]).Value != null)
                        {
                            desc = (e.Properties["description"]).Value.ToString();
                        }
                        else
                        {
                            desc = " ";
                        }
                        String login = (e.Properties["sAMAccountName"]).Value.ToString();
                        if (desc != null)
                        {
                            Person p2 = new JavaScriptSerializer().Deserialize<Person>(desc.ToString());
                            if (p2 == null)
                            {
                                Console.Write("{0} {1} {2}", desc.ToString(), " ", " ");
                                HTML_to_doc(FIO, login, " ", " ", mail, "");
                            }
                            else
                            {
                                Console.Write("{0} {1} {2}", desc.ToString(), p2.PAS, p2.SKD);
                                HTML_to_doc(FIO, login, p2.PAS, p2.SKD.ToString(), mail, "");
                            }
                        }
                    }
                    if (yy == Convert.ToInt32(nomer))
                    {
                        var e = (DirectoryEntry)p.GetUnderlyingObject();
                        String FIO = (e.Properties["cn"]).Value.ToString();
                        String mail = (e.Properties["mail"]).Value.ToString();
                        String desc;
                        if ((e.Properties["description"]) != null)
                        {
                            desc = (e.Properties["description"]).Value.ToString();
                        }
                        else
                        {
                            desc = " ";
                        }
                        String login = (e.Properties["sAMAccountName"]).Value.ToString();
                        if (desc != null)
                        {
                            Person p2 = new JavaScriptSerializer().Deserialize<Person>(desc.ToString());
                            Console.Write("{0} {1} {2}", desc.ToString(), p2.PAS, p2.SKD);
                            HTML_to_doc(FIO, login, p2.PAS, p2.SKD.ToString(), mail, "");
                        }
                    }
                    yy++;
                }


                grp.Dispose();
                ctx.Dispose();
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine("\nWe did not find that group in that domain, perhaps the group resides in a different domain?");
                Console.ReadLine();
            }
        }


        public void GetADUserInfo()
        {
            try
            {
                //This is a generic LDAP call, it would do a DNS lookup to find a DC in your AD site, scales better
                DirectoryEntry enTry = new DirectoryEntry("LDAP://OU=MyUsers,DC=Steve,DC=Schofield,DC=com");

                DirectorySearcher mySearcher = new DirectorySearcher(enTry);
                mySearcher.Filter = "(&(objectClass=user)(anr=smith))";

                try
                {
                    foreach (SearchResult resEnt in mySearcher.FindAll())
                    {
                        var DE = resEnt.GetDirectoryEntry();
                        var PP = DE.Properties;
                        var tt = PP.PropertyNames;
                    }
                }
                catch (Exception f)
                {
                    Console.WriteLine(f.Message);
                }
            }
            catch (Exception f)
            {
                Console.WriteLine(f.Message);
            }
        }
    }
}