using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Programm_Schichtpläne
{
   
    public partial class Form1 : Form
    {
        Timer timer = new Timer();
        public static DateTime Now { get; } // Uhrzeit holen
        public string t = "";
        public string t1 = "";
       
        public Form1()
        {
            InitializeComponent();
            timer.Interval = 1000; // Setzt das Intervall auf 1 Sekunde (1000 Millisekunden)
            timer.Tick += new EventHandler(timer1_Tick); // Fügt das Event hinzu, das bei jedem Tick aufgerufen wird
            timer.Start(); // Startet den Time
            
        }
        public static DateTime GetQuarterEnd(DateTime date)
        {
            int quarterNumber = (date.Month - 1) / 3+1;
            return new DateTime(date.Year, quarterNumber * 3, DateTime.DaysInMonth(date.Year, quarterNumber * 3));
        }
        public void druck(string pfad)
        {
            // Erstellen Sie eine neue Anwendung.
            Excel.Application excelApp = new Excel.Application();

            // Öffnen Sie die Excel-Datei.
            Excel.Workbook workbook = excelApp.Workbooks.Open(pfad);

            // Drucken Sie das gesamte Arbeitsbuch aus.
            workbook.PrintOutEx();

            // Schließen Sie das Arbeitsbuch und die Anwendung.
            workbook.Close();
            excelApp.Quit();
        }
        public void abfragefrüMO(string pfad)
        {
            // Abfrage der frühschicht Montag 
            if ((t == "Di") || (t == "Mi") || (t == "Do") || (t == "Fr"))
            {
                pfad = @"ihrpfad";
                druck(pfad);
                pfad = "";
            }
            else if (t == "Mo")
            {
               pfad = @"ihrpfad";
                druck(pfad);
                pfad = "";
            }
            else if (t == "Sa")
            {
                if (radioButton1.Checked == true)
                {
                    pfad = @"ihrpfad";
                    druck(pfad);
                    pfad = "";
                }
                pfad = @"ihrpfad";
                druck(pfad);
                pfad = "";
            }


        }
        private void btnDruck_Click(object sender, EventArgs e)
        {
            DateTime date = DateTime.Now; // Setzen Sie hier Ihr Datum
            DateTime quarterEnd = GetQuarterEnd(date);

            //Datum letzter Tag des Monats
            DateTime letzterTagDesMonats = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));


            string pfad = "";
           
            if (comboBox1.SelectedIndex == 0)
            {
                 if (t1 == "01")
                 {
                     pfad = @"L:\\operator\\SCHICHT\\4.    Sonstige\\1.    Erste Schicht im Monat.xlsx";
                     druck(pfad);
                     pfad = "";
                     abfragefrüMO(pfad);
                 }
                 // Überprüfen Sie, ob der aktuelle Tag der 30. ist, ob heute Samstag ist und ob das Quartalsende auf den nächsten Tag fällt
                 else if ((Convert.ToInt16(t1) + 1).ToString() == "31" && t == "Sa" && date.Date == quarterEnd.AddDays(-1))
                 {
                     // Wenn alle Bedingungen erfüllt sind, führen Sie die folgenden Aktionen aus
                     pfad = @"L:\\operator\\SCHICHT\\4.    Sonstige\\3.    Quartalsende.xlsx";
                     druck(pfad);
                     pfad = "";
               
                     pfad = @"L:\\operator\\SCHICHT\\4.    Sonstige\\2.    Letzte Schicht im Monat.xlsx";
                     druck(pfad);
                     pfad = "";
                 }
                 else
                 {
                     abfragefrüMO(pfad);
                 }
            }
            else if (comboBox1.SelectedIndex == 1)
            {
            // Spätschicht abfrage
                if ((t == "Mo") || (t == "Di") || (t == "Mi") || (t == "Do") || (t == "Fr"))
                {
                    pfad = @"ihrpfad";
                    druck(pfad);
                    pfad = "";
                }
                else if (t == "Sa")
                {
                   pfad = @"ihrpfad";
                    druck(pfad);
                    pfad = "";
                }
            }
            else if (comboBox1.SelectedIndex == 2)
            {
            // Abfrage Nachtschicht
                if ((t == "Mo") || (t == "Di") || (t == "Mi") || (t == "Do"))
                {
                    pfad = @"ihrpfad";
                    druck(pfad);
                    pfad = "";
                    //Quartals Ende ABfrage 
                    if (date.Date == quarterEnd)
                    {
                        pfad = @"ihrpfad";
                        druck(pfad);
                        pfad = "";

                    }
                    // Letzter Tag des Monats (Nachtschicht)
                    if (date.Date == letzterTagDesMonats)
                    {
                        pfad = @"ihrpfad";
                        druck(pfad);
                        pfad = "";
                    }
                }
                // Nachtschicht Freitag
                else if (t == "Fr")
                {
                    pfad = @"ihrpfad";
                    druck(pfad);
                    pfad = "";
                    //Quartals Ende ABfrage 
                    if (date.Date == quarterEnd)
                    {
                        pfad = @"ihrpfad";
                        druck(pfad);
                        pfad = "";
                    }
                    // Letzter Tag des Monats (Nachtschicht)
                    if (date.Date == letzterTagDesMonats)
                    {
                        pfad = @"ihrpfad";
                        druck(pfad);
                        pfad = "";
                    }
                }
                // Nachtschicht Sonntag
                else if(t== "So")
                {
                    pfad = @"ihrpfad";
                    druck(pfad);
                    pfad = "";
                  //Quartals Ende ABfrage 
                    if (date.Date == quarterEnd)
                    {
                        pfad = @"ihrpfad";
                        druck(pfad);
                        pfad = "";

                    }
                    // Letzter Tag des Monats (Nachtschicht)
                    else if (date.Date == letzterTagDesMonats)
                    {
                       pfad = @"ihrpfad";
                        druck(pfad);
                        pfad = "";
                    }
                 
                }
                // Nachtschicht Montag
                else if (t == "Mo")
                {
                   // BUtton pressed Feiertag
                    if (radioButton2.Checked == true) // Nachtschicht Montag Feiertag
                    {
                        pfad = @"ihrpfad";
                        druck(pfad);
                        pfad = "";
                        //Quartals Ende ABfrage 
                        if (date.Date == quarterEnd)
                        {
                           pfad = @"ihrpfad";
                            druck(pfad);
                            pfad = "";
                        }
                        // Letzter Tag des Monats (Nachtschicht)
                        if (date.Date == letzterTagDesMonats)
                        {
                            pfad = @"ihrpfad";
                            druck(pfad);
                            pfad = "";
                        }
                    }
                }
            }
            // Feiertags abfrage
            if (radioButton2.Checked == true)
            {
                pfad = @"ihrpfad";
                druck(pfad);
                pfad = "";
            }

            
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
        // Timer werte als Datum zu bekommen t = ddd = Mo,Di,Mi, als zahl dann t1 = dd = 01,02,03 usw.
            t = DateTime.Now.ToString("ddd");
            t1= DateTime.Now.ToString("dd");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
          
        }

      

        private void radioButton2_Click(object sender, EventArgs e)
        {
           
        }
    }
}
