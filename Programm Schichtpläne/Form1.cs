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
        public string t = ""; //deklarien
        public string t1 = ""; //deklarien
       
        public Form1()
        {
            InitializeComponent();
            timer.Interval = 1000; // Setzt das Intervall auf 1 Sekunde (1000 Millisekunden)
            timer.Tick += new EventHandler(timer1_Tick); // Fügt das Event hinzu, das bei jedem Tick aufgerufen wird
            timer.Start(); // Startet den Time
            
        }
        public static DateTime GetQuarterEnd(DateTime date)
        {
            int quarterNumber = (date.Month - 1) / 3+1; // Quartalsende
            return new DateTime(date.Year, quarterNumber * 3, DateTime.DaysInMonth(date.Year, quarterNumber * 3)); // return Quartalsende
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

            DateTime letzterTagDesMonats = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));


            string pfad = "";
            if (comboBox1.SelectedIndex == 0)
            {
                if (t1 == "01")
                {
                    pfad = @"ihrpfad";
                    druck(pfad);
                    pfad = "";
                    abfragefrüMO(pfad);
                }
                else
                {
                    abfragefrüMO(pfad);
                }
            }
            else if (comboBox1.SelectedIndex == 1)
            {
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
                if ((t == "Mo") || (t == "Di") || (t == "Mi") || (t == "Do"))
                {
                    pfad = @"ihrpfad";
                    druck(pfad);
                    pfad = "";
                    if (date.Date == quarterEnd)
                    {
                        pfad = @"ihrpfad";
                        druck(pfad);
                        pfad = "";

                    }
                    if (date.Date == letzterTagDesMonats)
                    {
                        pfad = @"ihrpfad";
                        druck(pfad);
                        pfad = "";
                    }
                }
                else if (t == "Fr")
                {
                    pfad = @"ihrpfad";
                    druck(pfad);
                    pfad = "";
                    if (date.Date == quarterEnd)
                    {
                        pfad = @"ihrpfad";
                        druck(pfad);
                        pfad = "";
                    }
                    if (date.Date == letzterTagDesMonats)
                    {
                        pfad = @"ihrpfad";
                        druck(pfad);
                        pfad = "";
                    }
                }
                else if (t == "Mo")
                {
                    if (radioButton1.Checked == true) // Nachtschicht Montag Feiertag
                    {
                        pfad = @"ihrpfad";
                        druck(pfad);
                        pfad = "";
                        if (date.Date == quarterEnd)
                        {
                            pfad = @"ihrpfad";
                            druck(pfad);
                            pfad = "";
                        }
                        if (date.Date == letzterTagDesMonats)
                        {
                            pfad = @"ihrpfad";
                            druck(pfad);
                            pfad = "";
                        }
                    }
                }
            }
            if (radioButton2.Checked == true)
            {
                pfad = @"ihrpfad";
                druck(pfad);
                pfad = "";
            }

            
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            t = DateTime.Now.ToString("ddd"); //t = Wochentag abkürzung also bsp. "Sa"
            t1= DateTime.Now.ToString("dd"); // t1 = ist der Tag also bsp. 28
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0) // Abfrage auswahl combobox
            {
                if (t == "Sa")
                {
                    radioButton1.Visible = true; // wenn Samstag ist wird der radiobutton sichtbar 
                }
                else
                {
                    radioButton1.Visible = false; // sonst nciht sichtbar
                }
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                if (t == "Mo")
                {
                    radioButton1.Visible = true; // wenn montag ist wird Feiertag und kommi wird der radio button auch sichtbar
                }
                {
                    radioButton1.Visible = false; // sonst nicht sichtbar
                }
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
           
        }

      

        private void radioButton2_Click(object sender, EventArgs e)
        {
           
        }
    }
}
