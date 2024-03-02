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
            
            if ((t == "Di") || (t == "Mi") || (t == "Do") || (t == "Fr"))
            {
                pfad = @"L:\\operator\\SCHICHT\\1.    Frühschicht\\2.    Frühschicht Dienstag - Freitag.xlsx";
                druck(pfad);
                pfad = "";
            }
            else if (t == "Mo")
            {
                pfad = @"L:\\operator\\SCHICHT\\1.    Frühschicht\\1.    Frühschicht Montag.xlsx";
                druck(pfad);
                pfad = "";
            }
            else if (t == "Sa")
            {
                if (radioButton1.Checked == true)
                {
                    pfad = @"L:\\operator\\SCHICHT\\1.    Frühschicht\\4.    Frühschicht Samstag Kommi.xlsx";
                    druck(pfad);
                    pfad = "";
                }
                pfad = @"L:\\operator\\SCHICHT\\1.    Frühschicht\\3.    Frühschicht Samstag.xlsx";
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
                    pfad = @"L:\\operator\\SCHICHT\\4.    Sonstige\\1.    Erste Schicht im Monat.xlsx";
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
                    pfad = @"L:\\operator\\SCHICHT\\2.    Spätschicht\\Spätschicht Montag - Freitag.xlsx";
                    druck(pfad);
                    pfad = "";
                }
                else if (t == "Sa")
                {
                    pfad = @"L:\\operator\\SCHICHT\\2.    Spätschicht\\Spätschicht Samstag.xlsx";
                    druck(pfad);
                    pfad = "";
                }
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                if ((t == "Mo") || (t == "Di") || (t == "Mi") || (t == "Do"))
                {
                    pfad = @"L:\\operator\\SCHICHT\\3.    Nachtschicht\\2.    Nachtschicht Montag - Donnerstag.xlsx";
                    druck(pfad);
                    pfad = "";
                    if (date.Date == quarterEnd)
                    {
                        pfad = @"L:\\operator\\SCHICHT\\4.    Sonstige\\3.    Quartalsende.xlsx";
                        druck(pfad);
                        pfad = "";

                    }
                    if (date.Date == letzterTagDesMonats)
                    {
                        pfad = @"L:\\operator\\SCHICHT\\4.    Sonstige\\2.    Letzte Schicht im Monat.xlsx";
                        druck(pfad);
                        pfad = "";
                    }
                }
                else if (t == "Fr")
                {
                    pfad = @"L:\\operator\\SCHICHT\\3.    Nachtschicht\\3.    Nachtschicht Freitag.xlsx";
                    druck(pfad);
                    pfad = "";
                    if (date.Date == quarterEnd)
                    {
                        pfad = @"L:\\operator\\SCHICHT\\4.    Sonstige\\3.    Quartalsende.xlsx";
                        druck(pfad);
                        pfad = "";
                    }
                    if (date.Date == letzterTagDesMonats)
                    {
                        pfad = @"L:\\operator\\SCHICHT\\4.    Sonstige\\2.    Letzte Schicht im Monat.xlsx";
                        druck(pfad);
                        pfad = "";
                    }
                }
                else if (t == "Mo")
                {
                    if (radioButton1.Checked == true) // Nachtschicht Montag Feiertag
                    {
                        pfad = @"L:\\operator\\SCHICHT\\3.    Nachtschicht\\4.    Nachtschicht Montag Feiertag Kommi.xlsx";
                        druck(pfad);
                        pfad = "";
                        if (date.Date == quarterEnd)
                        {
                            pfad = @"L:\\operator\\SCHICHT\\4.    Sonstige\\3.    Quartalsende.xlsx";
                            druck(pfad);
                            pfad = "";
                        }
                        if (date.Date == letzterTagDesMonats)
                        {
                            pfad = @"L:\\operator\\SCHICHT\\4.    Sonstige\\2.    Letzte Schicht im Monat.xlsx";
                            druck(pfad);
                            pfad = "";
                        }
                    }
                }
            }
            if (radioButton2.Checked == true)
            {
                pfad = @"L:\\operator\\SCHICHT\\4.    Sonstige\\5.    Feiertagsarbeit.xlsx";
                druck(pfad);
                pfad = "";
            }

            
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            t = DateTime.Now.ToString("ddd");
            t1= DateTime.Now.ToString("dd");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                if (t == "Sa")
                {
                    radioButton1.Visible = true;
                }
                radioButton1.Visible = false;
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                if (t == "Mo")
                {
                    radioButton1.Visible = true;
                }
                radioButton1.Visible = false;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
           if(radioButton2.Checked == true)
            {
                radioButton1.Checked = false;
            }
            else
            {
                radioButton2.Checked = true;
            }
        }

      

        private void radioButton2_Click(object sender, EventArgs e)
        {
           
        }
    }
}
