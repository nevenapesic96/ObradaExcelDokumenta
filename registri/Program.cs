using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.IO;
using MySql.Data.MySqlClient;


namespace registri
{
    class Evidencija
    {
        public String naziv = null;
        public DateTime? datumUspostavljanja = null;
        public DateTime? datumPrestanka = null;
        public String rokCuvanjaPodataka = null;
        public String vodjenjeEvidencije = null;
        public String nacinUnosaPodataka = null;
        public String periodAzuriranjaPodataka = null;
        public String centralizovanostPodataka = null;
        public String evidencijaFizickiSmestena = null;
        public String ITPodrska = null;
        public String pristupPomocuWebServisa = null;
        public String pristupDrugihOrganaPreko = null;
        public String pristupPodacima = null;
        public String potrebniPodaciDrugihOrgana = null;

        public int? propis;
        public String poslovi = null;
        public int? organ;


        public void popuniEvidenciju(Worksheet excelSheet, MySqlCommand command)
        {

            String nazivOrgana = excelSheet.Cells[2, 2].Value.ToString();
            organ = Program.nadjiOrgan(nazivOrgana,command);
            

            try
            {
                naziv = excelSheet.Cells[2, 4].Value.ToString();
               // Console.WriteLine(naziv);
            }
            catch(Exception)
            {
                Console.WriteLine("Greska prilikom upisa naziva evidencije");
                naziv = null;
            }

            try
            {
                /*String datum = excelSheet.Cells[2, 5].Value.ToString();

                if (datum == "*")
                    datumUspostavljanja = null;
                else
                    datumUspostavljanja = datum;

               // datumUspostavljanja = DateTime.ParseExact(datum, "dd.MM.yyyy.", System.Globalization.CultureInfo.InvariantCulture);
              //  String formattedDate = datumUspostavljanja.ToString("yyyy-MM-dd");

                Console.WriteLine(datumUspostavljanja);*/


            }
            catch (Exception)
            {
                Console.WriteLine("Greska prilikom upisa datumUs evidencije");
                 datumUspostavljanja=null;
            }

            try
            {
                /*
                 String datum=excelSheet.Cells[2, 6].Value.ToString();

                 if (datum == "*")
                     datumPrestanka = null;
                 else
                     datumPrestanka = datum;
               //     datumPrestanka= DateTime.ParseExact(datum, "dd.MM.yyyy.", null);

                Console.WriteLine(datumPrestanka); */


            }
            catch (Exception)
            {
                Console.WriteLine("Greska prilikom upisa datum p evidencije");
                datumPrestanka=null;
            }

            try
            {
                String p=excelSheet.Cells[2, 7].Value.ToString();
                propis = Program.nadjiPropis(p,command);
              //  propis = Int32.Parse(p);

                //Console.WriteLine(propis);
                
            }
            catch (Exception)
            {
                Console.WriteLine("Greska prilikom upisa propis evidencije");
                propis=null;
            }

            try
            {

                vodjenjeEvidencije = excelSheet.Cells[2, 8].Value.ToString();
                if(vodjenjeEvidencije=="*")
                    vodjenjeEvidencije=null;

               // Console.WriteLine(vodjenjeEvidencije);
            }
            catch (Exception)
            {
                Console.WriteLine("Greska prilikom upisa vodjenje evidencije");
                vodjenjeEvidencije=null;
            }

            try
            {

                rokCuvanjaPodataka = excelSheet.Cells[2, 9].Value2.ToString();

               //  Console.WriteLine(rokCuvanjaPodataka);
            }
            catch (Exception)
            {
                Console.WriteLine("Greska prilikom upisa rok cuvanja podataka");
                rokCuvanjaPodataka = null;
            }

            try
            {

                nacinUnosaPodataka = excelSheet.Cells[2, 10].Value2.ToString();
                if(nacinUnosaPodataka=="*")
                    nacinUnosaPodataka=null;

              //   Console.WriteLine(nacinUnosaPodataka);
            }
            catch (Exception)
            {
                Console.WriteLine("Greska prilikom upisa nacin unosa evidencije");
                nacinUnosaPodataka=null;
            }

            try
            {

                periodAzuriranjaPodataka = excelSheet.Cells[2, 12].Value2.ToString();
                if(periodAzuriranjaPodataka=="*")
                    periodAzuriranjaPodataka=null;

               //  Console.WriteLine(periodAzuriranjaPodataka);
            }
            catch (Exception)
            {
                Console.WriteLine("Greska prilikom upisa period a evidencije");
                periodAzuriranjaPodataka=null;
            }

            try
            {

                centralizovanostPodataka = excelSheet.Cells[2, 14].Value2.ToString();
                if(centralizovanostPodataka=="*")
                    centralizovanostPodataka=null;

              //   Console.WriteLine(centralizovanostPodataka);
            }
            catch (Exception)
            {
                Console.WriteLine("Greska prilikom upisa centr evidencije");
                centralizovanostPodataka=null;
            }

            try
            {

                evidencijaFizickiSmestena = excelSheet.Cells[2, 16].Value2.ToString();
                if (evidencijaFizickiSmestena == "*")
                    evidencijaFizickiSmestena = null;

               //  Console.WriteLine(evidencijaFizickiSmestena);
            }
            catch (Exception)
            {
                Console.WriteLine("Greska prilikom upisa ev fiz s evidencije");
                evidencijaFizickiSmestena = null;
            }

            try
            {

                ITPodrska = excelSheet.Cells[2, 17].Value2.ToString();
                if (ITPodrska == "*")
                    ITPodrska = null;

               //  Console.WriteLine(ITPodrska);
            }
            catch (Exception)
            {
                Console.WriteLine("Greska prilikom upisa it evidencije");
                ITPodrska = null;
            }

            try
            {

                pristupPomocuWebServisa = excelSheet.Cells[2, 18].Value2.ToString();
                if (pristupPomocuWebServisa == "*")
                    pristupPomocuWebServisa = null;

               //  Console.WriteLine(pristupPomocuWebServisa);
            }
            catch (Exception)
            {
                Console.WriteLine("Greska prilikom upisa web s evidencije");
                pristupPomocuWebServisa = null;
            }

            try
            {

                pristupDrugihOrganaPreko = excelSheet.Cells[2, 19].Value2.ToString();
                if (pristupDrugihOrganaPreko == "*")
                    pristupDrugihOrganaPreko = null;

               //  Console.WriteLine(pristupDrugihOrganaPreko);
            }
            catch (Exception)
            {
                pristupDrugihOrganaPreko = null;
                Console.WriteLine("Greska prilikom upisa pristup dr org evidencije");
            }

            try
            {

                pristupPodacima = excelSheet.Cells[2, 20].Value2.ToString();
                if (pristupPodacima == "*")
                    pristupPodacima = null;

                // Console.WriteLine(pristupPodacima);
            }
            catch (Exception)
            {
                Console.WriteLine("Greska prilikom upisa pristup p evidencije");
                pristupPodacima = null;
            }

            try
            {

                potrebniPodaciDrugihOrgana = excelSheet.Cells[2, 22].Value2.ToString();
                if(potrebniPodaciDrugihOrgana=="*")
                    potrebniPodaciDrugihOrgana=null;

              //   Console.WriteLine(potrebniPodaciDrugihOrgana);
            }
            catch (Exception)
            {
                Console.WriteLine("Greska prilikom upisa pristup drugih org evidencije");
                potrebniPodaciDrugihOrgana = null;
            }
        }

    }


    class Program
    {
        //ovo ne treba, samo test kako radi fja
        static public void ucitaj()
        {
            string[] fajlovi = Directory.GetFiles("C:\\Users\\NEVEN\\Desktop\\buk");

            Application excel = new Application();

            foreach (string fajl in fajlovi)
            {
                Workbook wb = excel.Workbooks.Open(fajl);
                Worksheet excelSheet = wb.ActiveSheet;

                var test = excelSheet.Cells[2, 4].Value.ToString();
               
                wb.Close();

                Console.WriteLine(test);
            }

        }


        //konekcija se pravi, pa se fja poziva svaki put kada se povezuje sa bazom
        static public MySqlConnection konektujSe()
        {
            try
            {
                string connection = "server=127.0.0.1; database=pois; user=root; password=; charset=utf8";
                MySqlConnection conn = new MySqlConnection(connection);
                conn.Open();
                return conn;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Konekcija nije uspela");
            }

            return null;

        }

        //popunjava tabelu podoblasti
        static public void ucitajPodoblasti()
        {
            //upisi pravu putanju do foldera gde su exceli
            string[] fajlovi = Directory.GetFiles("C:\\Users\\Vukasin\\Downloads\\podoblasti");
            Application excel = new Application();
            LinkedList<String> podoblasti = new LinkedList<string>();

         //   int brfajla = 0;
            int i;
            int id = 1;

            foreach (string fajl in fajlovi)
            {
                i = 3;

                Workbook wb = excel.Workbooks.Open(fajl);
                //upisi naziv sheeta
                Worksheet worksheet = (Worksheet)wb.Worksheets["Sheet1"];

                MySqlConnection conn = konektujSe();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;


                String test = worksheet.Cells[1, 1].Value.ToString();
                int oblastID = nadjiOblast(test, command);
                Console.WriteLine(oblastID);

                
                while (i != 0)
                {

                    String podoblast = "";
                    try
                    {
                        podoblast = worksheet.Cells[i, 2].Value.ToString();
                    }
                    catch (Exception)
                    {
                  //      i = 0;
                    }
                    //ispod poslednje podoblasti upisi obavezno "kraj"
                    if (podoblast == "kraj" || i == 0)
                        i = 0;
                    else
                    {
                        if (!podoblasti.Contains(podoblast))
                        {
                            command.CommandText = String.Format("INSERT INTO podoblast VALUES ({0},'{1}',{2})", id, podoblast, oblastID); ;
                            command.ExecuteNonQuery();
                            id++;
                            podoblasti.AddLast(podoblast);
                           // Console.WriteLine(podoblast);
                        }
                        i++;

                    }
                }


                conn.Close();
                wb.Close();

               // brfajla++;
               
            }

            excel.Quit();

        }

        //pomocna za gornju fju
        public static int nadjiOblast(String test, MySqlCommand command)
        {
            command.CommandText = String.Format("SELECT oblastID FROM oblast WHERE naziv='{0}'", test);
            MySqlDataReader reader = command.ExecuteReader();
            reader.Read();

            int id=Int32.Parse(reader["oblastID"].ToString());
            reader.Close();
            return (id);
        }

        //popunjava bazu pravni osnov
        static public void unesiPravniOsnov()
        {
            //upisi pravu putanju do excela
            String fajl="C:\\Users\\NEVEN\\Desktop\\buk\\Spisak evidencija";

             Application excel = new Application();

             Workbook wb = excel.Workbooks.Open(fajl);
            //upisi naziv sheeta
             Worksheet worksheet = (Worksheet) wb.Worksheets["pravni osnov"];
           
            MySqlConnection conn=konektujSe();
             MySqlCommand command = new MySqlCommand();
             command.Connection = conn;

             command.CommandText = "delete from pravniosnov" ;
             command.ExecuteNonQuery();
               
                int i = 3;
                while(i!=0)
            {

                String test = "";
                try
                {
                    test = worksheet.Cells[i, 1].Value.ToString();
                }
                catch (Exception)
                {
                    i = 0;
                }
                    //ispod poslednjeg propisa upisi obavezno "kraj"
                     if (test == "kraj" || i==0)
                    i = 0;
                     else
                {
                   
                   command.CommandText = String.Format("INSERT INTO pravniosnov VALUES ({0},'{1}','{2}')",i-2, test, null); ;
                   command.ExecuteNonQuery();
                   Console.WriteLine(i);
                  
                   i++;
                }
			}
                
            conn.Close();
            wb.Close();
            excel.Quit();

        }

        public static int? nadjiOrgan(string naziv,MySqlCommand command)
        {
            command.CommandText = String.Format("SELECT organID FROM organ WHERE naziv='{0}'", naziv);
            MySqlDataReader reader = command.ExecuteReader();
            if (reader.Read())
            {
                int id;
                if (int.TryParse(reader["organID"].ToString(), out id))
                {
                    reader.Close();
                    return id;
                }
            }
            reader.Close();
            return null;
        }

        static public int? nadjiPropis(String p, MySqlCommand command)
        {

            command.CommandText = String.Format("SELECT propisID FROM propis WHERE naziv='{0}'", p);
            MySqlDataReader reader = command.ExecuteReader();
            if (reader.Read())
            {

                int id;
                if (int.TryParse(reader["propisID"].ToString(), out id))
                {
                    reader.Close();
                    return id;
                }
            }
            reader.Close();
            return null;

                     
        }

        //popunjava poslove, nisam proverila da li radi, moraju prvo podoblasti da se popune
        static public void ucitajPoslove()
        {

            //upisi pravu putanju do excela
            string[] fajlovi = Directory.GetFiles("C:\\Users\\Vukasin\\Downloads\\podoblasti");
            Application excel = new Application();


            int brfajla = 0;
            int i;
            int posloviID = 1;

            foreach (string fajl in fajlovi)
            {
                i = 3;

                Workbook wb = excel.Workbooks.Open(fajl);
                //upisi naziv sheeta
                Worksheet worksheet = (Worksheet)wb.Worksheets["Sheet1"];

                MySqlConnection conn = konektujSe();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                Console.WriteLine(brfajla+1);

                
                while (i != 0)
                {
                    String test="";
                    bool predji = false;
                     int podOblastID=0;

                     try
                     {
                         test = worksheet.Cells[i, 2].Value.ToString();

                     }
                     catch (Exception)
                     {
                         podOblastID = 1;
                     }
                    try
                    {
                        if(podOblastID!=1)
                         podOblastID= nadjiPodOblast(test, command);
                    }
                    catch (Exception)
                    {
                        test = "kraj";
                    }
                   

                    String opis = "";
                    try
                    {
                        opis = worksheet.Cells[i, 3].Value.ToString();
                        
                    }
                    catch (Exception)
                    {
                    //    i = 0;
                        predji = true;
                    }
                    //ispod poslednje podoblasti upisi obavezno "kraj"
                    if (test == "kraj" || i == 0)
                        i = 0;
                    else
                    {
                        if (!predji)
                        {
                            command.CommandText = String.Format("INSERT INTO poslovi VALUES ({0},'{3}','{1}',{2})", posloviID, opis, podOblastID, null); ;
                            command.ExecuteNonQuery();
                            posloviID++;
                           
                        }
                        i++;
                    }
                }


                conn.Close();
                wb.Close();

                brfajla++;

            }

            excel.Quit();

        }

        //pomocna fja, za gornju
        public static int nadjiPodOblast(String test, MySqlCommand command)
        {
            try
            {
                command.CommandText = String.Format("SELECT podoblastID FROM podoblast WHERE naziv='{0}'", test);
                MySqlDataReader reader = command.ExecuteReader();
                reader.Read();

                int id = Int32.Parse(reader["podoblastID"].ToString());
                reader.Close();
                return (id);
            }
            catch (Exception)
            {
                return 1;
            }
        }

        public static void ucitajPropise()
        {

            String fajl = "C:\\Users\\NEVEN\\Desktop\\buk\\Spisak evidencija";

            Application excel = new Application();

            Workbook wb = excel.Workbooks.Open(fajl);
            Worksheet worksheet = (Worksheet)wb.Worksheets["pravni osnov"];

            MySqlConnection conn = konektujSe();
            MySqlCommand command = new MySqlCommand();
            command.Connection = conn;

            command.CommandText = "delete from propis" ;
            command.ExecuteNonQuery();

            int i = 3;
            while (i != 0)
            {

                String test = "";
                try
                {
                    test = worksheet.Cells[i, 1].Value.ToString();
                }
                catch (Exception)
                {
                    i = 0;
                }
                //ispod poslednjeg propisa upisi obavezno "kraj"
                if (test == "kraj" || i == 0)
                    i = 0;
                else
                {

                    command.CommandText = String.Format("INSERT INTO propis VALUES ({0},'{1}','{2}',null,{4});", i - 2, test, null, null, i-2);
                    command.ExecuteNonQuery();
                    Console.WriteLine(i);

                    i++;
                }
            }

            conn.Close();
            wb.Close();
            excel.Quit();

        }


        static public void evidencije()
        {
            string[] fajlovi = Directory.GetFiles("C:\\Users\\NEVEN\\Desktop\\ZIPOVANEEVIDENCIJE");

            Application excel = new Application();

            MySqlConnection conn = konektujSe();
            MySqlCommand command = new MySqlCommand();
            command.Connection = conn;

            command.CommandText = "delete from evidencija";
            command.ExecuteNonQuery();

            var sb = new StringBuilder();
            String fajlT = null;

            int i = 1;
            foreach (string fajl in fajlovi)
            {
                Workbook wb = excel.Workbooks.Open(fajl);
                Worksheet excelSheet = wb.ActiveSheet;

                Evidencija evidencija = new Evidencija();
                try
                {
                evidencija.popuniEvidenciju(excelSheet,command);

                
                    if (evidencija.organ == null)
                    {
                        command.CommandText = String.Format("INSERT INTO evidencija(evidencijaID,naziv,rokCuvanjaPodataka,vodjenjeEvidencije,nacinUnosaPodataka,periodAzuriranjaPodataka,centralizovanostPodataka,evidencijaFizickiSmestena,ITPodrska,pristupPomocuWebServisa,pristupDrugihOrganaPreko,pristupPodacima,potrebniPodaciDrugihOrgana,propisID,organID)" +
                    "VALUES (DEFAULT,'{0}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}',{14},NULL)",
                       evidencija.naziv, evidencija.datumUspostavljanja, evidencija.datumPrestanka, evidencija.rokCuvanjaPodataka, evidencija.vodjenjeEvidencije, evidencija.nacinUnosaPodataka, evidencija.periodAzuriranjaPodataka, evidencija.centralizovanostPodataka, evidencija.evidencijaFizickiSmestena, evidencija.ITPodrska, evidencija.pristupPomocuWebServisa, evidencija.pristupDrugihOrganaPreko, evidencija.pristupPodacima, evidencija.potrebniPodaciDrugihOrgana, evidencija.propis, evidencija.organ);
                    }
                    else
                    {
                        if (evidencija.propis == null)
                        {
                            command.CommandText = String.Format("INSERT INTO evidencija(evidencijaID,naziv,rokCuvanjaPodataka,vodjenjeEvidencije,nacinUnosaPodataka,periodAzuriranjaPodataka,centralizovanostPodataka,evidencijaFizickiSmestena,ITPodrska,pristupPomocuWebServisa,pristupDrugihOrganaPreko,pristupPodacima,potrebniPodaciDrugihOrgana,propisID,organID)" +
                "VALUES (DEFAULT,'{0}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}',NULL,{15})",
                   evidencija.naziv, evidencija.datumUspostavljanja, evidencija.datumPrestanka, evidencija.rokCuvanjaPodataka, evidencija.vodjenjeEvidencije, evidencija.nacinUnosaPodataka, evidencija.periodAzuriranjaPodataka, evidencija.centralizovanostPodataka, evidencija.evidencijaFizickiSmestena, evidencija.ITPodrska, evidencija.pristupPomocuWebServisa, evidencija.pristupDrugihOrganaPreko, evidencija.pristupPodacima, evidencija.potrebniPodaciDrugihOrgana, evidencija.propis, evidencija.organ);
                        }
                        else
                        {
                            command.CommandText = String.Format("INSERT INTO evidencija(evidencijaID,naziv,rokCuvanjaPodataka,vodjenjeEvidencije,nacinUnosaPodataka,periodAzuriranjaPodataka,centralizovanostPodataka,evidencijaFizickiSmestena,ITPodrska,pristupPomocuWebServisa,pristupDrugihOrganaPreko,pristupPodacima,potrebniPodaciDrugihOrgana,propisID,organID)" +
                   "VALUES (DEFAULT,'{0}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}',{14},{15})",
                      evidencija.naziv, evidencija.datumUspostavljanja, evidencija.datumPrestanka, evidencija.rokCuvanjaPodataka, evidencija.vodjenjeEvidencije, evidencija.nacinUnosaPodataka, evidencija.periodAzuriranjaPodataka, evidencija.centralizovanostPodataka, evidencija.evidencijaFizickiSmestena, evidencija.ITPodrska, evidencija.pristupPomocuWebServisa, evidencija.pristupDrugihOrganaPreko, evidencija.pristupPodacima, evidencija.potrebniPodaciDrugihOrgana, evidencija.propis, evidencija.organ);
                        }
                    }
                    command.ExecuteNonQuery();


                    sb.AppendLine(fajl);
                    File.WriteAllText("C:\\Users\\NEVEN\\Desktop\\uneseneEv.txt", sb.ToString());

                    wb.Close(false);
                    Console.WriteLine(i);
                    i++;
                }catch(Exception)
                {
                    sb.AppendLine("********ovaj nece "+fajl);
                    File.WriteAllText("C:\\Users\\NEVEN\\Desktop\\uneseneEv.txt", sb.ToString());
                }

            }
            conn.Close();
            excel.Quit();
        }

        static public void procitajE()
        {
            //ovde izmeni putanju
            string[] fajlovi = Directory.GetFiles("C:\\Users\\NEVEN\\Desktop\\KONACNA VERZIJA EVIDENCIJA");

            Application excel = new Application();

            var sb = new StringBuilder();
            String fajlT = null;
            int i = 1;
            try
            {
                foreach (string fajl in fajlovi)
                {

                    fajlT = fajl;

                        Workbook wb = excel.Workbooks.Open(fajl);
                        Worksheet excelSheet = wb.ActiveSheet;

                        String naziv = excelSheet.Cells[2, 4].Value.ToString();

                        sb.AppendLine(naziv);

                        Console.WriteLine(i++);
                        wb.Close(false);

                }
            }catch(Exception){
                Console.WriteLine("greska kod "+i);
                sb.AppendLine("----------------"+fajlT);

            }


            File.WriteAllText("C:\\Users\\NEVEN\\Desktop\\spisakEv.txt", sb.ToString());

        }


        static public void preimenujE()
        {
            string[] fajlovi = Directory.GetFiles("C:\\Users\\NEVEN\\Desktop\\kopija");

            Application excel = new Application();

            var sb = new StringBuilder();
            int i = 1;
            foreach (string fajl in fajlovi)
            {
                Workbook wb = excel.Workbooks.Open(fajl);
                Worksheet excelSheet = wb.ActiveSheet;

                String naziv = excelSheet.Cells[2, 4].Value.ToString();

                wb.Close(false);

                try
                {
                    File.Move(fajl, "C:\\Users\\NEVEN\\Desktop\\ispravljene\\" + naziv + ".xlsx");
                }
                catch (IOException)
                {
                    Console.WriteLine("IOException");
                    File.Move(fajl, @"C:\Users\NEVEN\Desktop\dupleEv\" + naziv + "(" + i + ")" + ".xlsx");
                    sb.AppendLine("duplo--------" + naziv);
                    i++;
                }
                catch (NotSupportedException)
                {
                    Console.WriteLine("NotSupportedEx");
                    sb.AppendLine("greska--------" + naziv);
                    File.WriteAllText(@"C:\Users\NEVEN\Desktop\ispravljene.txt", sb.ToString());
                }
                catch (Exception)
                {
                    Console.WriteLine("Exception");
                }
                
            }
            File.WriteAllText("C:\\Users\\NEVEN\\Desktop\\ispravljene.txt", sb.ToString());

        }
       
        static public void ucitajOrgane()
        {
            String fajl = "C:\\Users\\NEVEN\\Desktop\\evidencija";

            Application excel = new Application();

            Workbook wb = excel.Workbooks.Open(fajl);
            Worksheet worksheet = (Worksheet)wb.Worksheets["Понуђени одговори"];

            MySqlConnection conn = konektujSe();
            MySqlCommand command = new MySqlCommand();
            command.Connection = conn;

            command.CommandText = "delete from organ";
            command.ExecuteNonQuery();

            int i = 2;
            while (i != 0)
            {
                int pib=-1;
                String test="";
                String naziv = "";
                try
                {
                    naziv = worksheet.Cells[i, 2].Value.ToString();
                    test = worksheet.Cells[i, 1].Value.ToString();
                    pib = Int32.Parse(test);
                                    }
                catch (Exception)
                {
                    i = 0;
                }
                //ispod poslednjeg propisa upisi obavezno "kraj"
                if (naziv == "kraj")
                    i = 0;
                else
                {

                    command.CommandText = String.Format("INSERT INTO organ VALUES ({0},'{1}',null,{3});", i - 1, naziv, null, pib);
                    command.ExecuteNonQuery();
                    Console.WriteLine(i);

                    i++;
                }
            }
            conn.Close();
            wb.Close();
            excel.Quit();
        }

        static public void ucitajPodatke()
        {
            //upisi pravu putanju do foldera gde su exceli
            string[] fajlovi = Directory.GetFiles("C:\\Users\\NEVEN\\Desktop\\ZIPOVANEEVIDENCIJE");
            Application excel = new Application();
            LinkedList<String> podaci = new LinkedList<string>();

            int i;

            MySqlConnection conn = konektujSe();
            MySqlCommand command = new MySqlCommand();
            command.Connection = conn;

            command.CommandText = "delete from organ";
            command.ExecuteNonQuery();


            foreach (string fajl in fajlovi)
            {
                i = 2;

                Workbook wb = excel.Workbooks.Open(fajl);
                //upisi naziv sheeta
                Worksheet worksheet = (Worksheet)wb.Worksheets["Упитник"];

                while (i != 0)
                {

                    String naziv = "";
                    try
                    {
                        naziv= worksheet.Cells[i, 13].Value.ToString();
                    }
                    catch (Exception)
                    {
                              i = 0;
                    }

                    //ispod poslednje podoblasti upisi obavezno "kraj"
                    if (i != 0)
                        
                    {
                        if (!podaci.Contains(naziv))
                        {
                            command.CommandText = String.Format("INSERT INTO podatak(podatakID,tip,naziv,upravniPostupakID) VALUES (DEFAULT,'','{0}',NULL)", naziv); ;
                            command.ExecuteNonQuery();
                            podaci.AddLast(naziv);
                            Console.WriteLine(fajl);
                        }
                        i++;
                    }
                }

                wb.Close();
            }

            conn.Close();
            excel.Quit();

        }

        public static int nadjiEvideniju(String naziv,MySqlCommand command)
        {
            try
            {
                command.CommandText = String.Format("SELECT evidencijaID FROM evidencija WHERE naziv='{0}'", naziv);
                MySqlDataReader reader = command.ExecuteReader();
                reader.Read();

                int id = Int32.Parse(reader["evidencijaID"].ToString());
                reader.Close();
                return (id);
            }
            catch (Exception)
            {
                return 1;
            }
        }

        public static int nadjiIDPodatak(String naziv,MySqlCommand command)
        {
            try
            {
                command.CommandText = String.Format("SELECT podatakID FROM podatak WHERE naziv='{0}'", naziv);
                MySqlDataReader reader = command.ExecuteReader();
                reader.Read();

                int id = Int32.Parse(reader["podatakID"].ToString());
                reader.Close();
                return (id);
            }
            catch (Exception)
            {
                return 1;
            }
        }

        public static void podatakUEvidenciji()
        {
            string[] fajlovi = Directory.GetFiles("C:\\Users\\NEVEN\\Desktop\\ZIPOVANEEVIDENCIJE");
            Application excel = new Application();

            int i;

            MySqlConnection conn = konektujSe();
            MySqlCommand command = new MySqlCommand();
            command.Connection = conn;

            command.CommandText = "delete from podatakUEvidenciji";
            command.ExecuteNonQuery();

            foreach (string fajl in fajlovi)
            {
                i = 2;
                
                Workbook wb = excel.Workbooks.Open(fajl);
                //upisi naziv sheeta
                Worksheet worksheet = (Worksheet)wb.Worksheets["Упитник"];

                String naziv=worksheet.Cells[2,4].Value.ToString();
                int idEvidencija=nadjiEvideniju(naziv,command);

                while(i!=0)
                {

                    String podatak="";
                    try
                    {
                        podatak = worksheet.Cells[i, 13].Value.ToString();
                    }
                    catch (Exception)
                    {
                        i = 0;
                    }

                    //ispod poslednje podoblasti upisi obavezno "kraj"
                    if (i != 0)
                    {
                            int idPodatak=nadjiIDPodatak(podatak,command);
                            command.CommandText = String.Format("INSERT INTO podatakUEvidenciji(id,evidencija,podatak) VALUES (DEFAULT,{0},{1})", idEvidencija,idPodatak); ;
                            command.ExecuteNonQuery();
                            Console.WriteLine(i);
                        }
                        i++;
                }

                wb.Close();
            }

            conn.Close();
            excel.Quit();
        }


        public static void ucitajSheetEvidencije(Workbook wb,Workbook krajnji, string fajl, int id)
        {
            Worksheet worksheet = (Worksheet)krajnji.Worksheets["evidencije"];
            Worksheet ws=(Worksheet)wb.Worksheets["Упитник"];
            Range last = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            int row = last.Row+1;

            try
            {
                worksheet.Cells[row, 1].Value = id;
            }
            catch (Exception)
            {
                Console.WriteLine(fajl+"  id");
            }
            
            
            for (int i = 1; i <= 12; i++)
            {
                try
                {
                    worksheet.Cells[row, i + 1].Value = ws.Cells[2, i].Value;
                }
                catch (Exception)
                {
                    Console.WriteLine(fajl+"  prva pol");
                }
            }
            for (int i = 14; i <= 21; i++)
            {
                try
                {
                    worksheet.Cells[row, i].Value = ws.Cells[2, i].Value;
                }
                catch (Exception)
                {
                    Console.WriteLine(fajl+"  druga pol");
                }
            }

            worksheet.Cells[row, 22].Value = fajl;
            krajnji.Save();

        }

        public static void ucitajSheetPodaci(Workbook wb, Workbook krajnji, string fajl, int id)
        {
            
        Worksheet wsPodaci = (Worksheet)krajnji.Worksheets["podaci"];
            Worksheet wsUpitnik = (Worksheet)wb.Worksheets["Упитник"];
            Range last = wsPodaci.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            int row = last.Row + 1;

            wsPodaci.Cells[row, 1].Value = id;
            wsPodaci.Cells[row, 2].Value = wsUpitnik.Cells[2, 4].Value;

            Boolean kraj=false;

            int i = 2;

            while (!kraj)
            {
                try{
                wsPodaci.Cells[row + i - 2, 3].Value = wsUpitnik.Cells[i, 13].Value.ToString();
                i++;

                }catch(Exception)
                {
                    kraj=true;
                }
            }

            krajnji.Save();
}
        public static void ucitajSheetDrugiOrganiPristup(Workbook wb, Workbook krajnji, string fajl, int id)
        {
            Worksheet wsOrgani = (Worksheet)krajnji.Worksheets["drugiOrganiPristup"];

            Worksheet wsUpitnik = (Worksheet)wb.Worksheets["Упитник"];

            Range last = wsOrgani.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            int row = last.Row + 1;
            
            Boolean kraj=false;

            int i = 5;

            while (!kraj)
            {
                try{
                    String prvo = wsUpitnik.Cells[i, 20].Value.ToString();
                    String drugo = wsUpitnik.Cells[i, 21].Value.ToString();
                    if (prvo.Equals("") || (drugo.Equals("")))
                    {
                        kraj = true;
                    }
                    else
                    {
                        
                        wsOrgani.Cells[row, 1].Value = id;
                        wsOrgani.Cells[row, 2].Value = wsUpitnik.Cells[2, 4].Value;
                        wsOrgani.Cells[row + i - 5, 3].Value = prvo;
                        wsOrgani.Cells[row + i - 5, 4].Value = drugo;

                            i++;
                    }

                }catch(Exception)
                {
                    kraj=true;
                }
            }

            krajnji.Save();
        }

        public static void ucitajSheetDrugiOrganiPodaci(Workbook wb, Workbook krajnji, string fajl, int id)
        {
            Worksheet wsOrgani = (Worksheet)krajnji.Worksheets["drugiOrganiPodaci"];

            Worksheet wsUpitnik = (Worksheet)wb.Worksheets["Упитник"];

            Range last = wsOrgani.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            int row = last.Row + 1;

            Boolean kraj = false;

            int i = 5;

            while (!kraj)
            {
                try
                {
                    String prvo = wsUpitnik.Cells[i, 22].Value.ToString();
                    String drugo = wsUpitnik.Cells[i, 23].Value.ToString();
                    if (prvo.Equals("") || (drugo.Equals("")))
                    {
                        kraj = true;
                    }
                    else
                    {
                        
                        wsOrgani.Cells[row, 1].Value = id;
                        wsOrgani.Cells[row, 2].Value = wsUpitnik.Cells[2, 4].Value;
                        wsOrgani.Cells[row + i - 5, 3].Value = prvo;
                        wsOrgani.Cells[row + i - 5, 4].Value = drugo;

                            i++;
                    }
                }
                catch (Exception)
                {
                    kraj = true;
                }
            }

            krajnji.Save();
        }
        public static void ubaciUJedanExcel()
        {
            string[] fajlovi = Directory.GetFiles("C:\\Users\\NEVEN\\Desktop\\KONACNAVERZIJAEVIDENCIJA");
            Application excel = new Application();
            int id=1;
            var sb = new StringBuilder();

            Workbook krajnji = excel.Workbooks.Open("C:\\Users\\NEVEN\\Desktop\\krajnjiExcel.xlsx");
            
            foreach (string fajl in fajlovi)
            {
                Workbook wb = null ;

                try
                {
                    wb = excel.Workbooks.Open(fajl);

                    ucitajSheetEvidencije(wb, krajnji, fajl, id);
                    ucitajSheetPodaci(wb, krajnji, fajl, id);
                    ucitajSheetDrugiOrganiPristup(wb, krajnji, fajl, id);
                    ucitajSheetDrugiOrganiPodaci(wb, krajnji, fajl, id);


                    Console.WriteLine(id);
                    sb.AppendLine(fajl);
                    wb.Close(false);
                }
                catch (Exception)
                {
                    sb.AppendLine("*****"+fajl);
                    Console.WriteLine(fajl + "    nece da otvori");
                    wb.Close(false);
                }


                id++;
            }
            krajnji.Close();
            excel.Quit();
            File.WriteAllText("C:\\Users\\NEVEN\\Desktop\\lose.txt", sb.ToString());

        }
        static void Main(string[] args)
        {

         //  unesiPravniOsnov();
          // ucitajPodoblasti();
          //  ucitajPoslove();
          //  ucitajPropise();
          //  evidencije();
            //procitajE();
           // preimenujE();
         //   ucitaj();
           // ucitajOrgane();
           // ucitajPodatke();
           // podatakUEvidenciji();
            ubaciUJedanExcel();
        }
    }
}
