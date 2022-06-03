// --------------------------------------------
// Datei:   Adrian Bicskei und Dion Elshani Lern- und Arbeitsauftrag Modul 403 1300
// Datum:   12. März 2020
// Ersteller: Adrian Bicskei & Dion Elshani, Berufsfachschule Baden
// Version: 1.3.0
// Änderungen:<12. März 2020/Adrian Bicskei/Hinzufügen von Programmkopf und Komentare>
// Beschreibung:
// Dies ist ein Programm, dass sich auf das Projektidee Wörtliabfragesystem basiert. Es werden also Wörter abgefragt, die aus einem anderem Datei stammen.
// --------------------------------------------


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Excel;       //Damit man mit Excel arbeiten kann

namespace ConsoleApp1
{
    class Class1
    {

        static int reihe = 2;       //Wert für Reihen, die sich ständig verändern
        static int punktZahl = 0;   //Wert für Punktzahl
        static int maxPunktZahl;    //Wert für Maximal erreichbare Punktzahl

        static void Main(string[] args)
        {
            anfang();
        }

        public static void anfang()
        {
            Console.WriteLine("\n\n\n");     //neue Zeilen   

            Console.ForegroundColor = ConsoleColor.Red;         //Französche und Deutsche Flagge
            Console.Write("\t█████████");
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("█████████");
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.Write("█████████");

            Console.ForegroundColor = ConsoleColor.Black;
            Console.Write("███████████████████████████");

            Console.Write("\n");

            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("\t█████████");
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("█████████");
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.Write("█████████");

            Console.ForegroundColor = ConsoleColor.Black;
            Console.Write("\t███████████████████████████");

            Console.Write("\n");

            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("\t█████████");
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("█████████");
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.Write("█████████");

            Console.ForegroundColor = ConsoleColor.Black;
            Console.Write("\t███████████████████████████");

            Console.Write("\n");

            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("\t█████████");
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("█████████");
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.Write("█████████");

            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("\t███████████████████████████");

            Console.Write("\n");

            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("\t█████████");
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("█████████");
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.Write("█████████");

            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("\t███████████████████████████");

            Console.Write("\n");

            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("\t█████████");
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("█████████");
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.Write("█████████");

            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("\t███████████████████████████");

            Console.Write("\n");

            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("\t█████████");
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("█████████");
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.Write("█████████");

            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.Write("\t███████████████████████████");

            Console.Write("\n");

            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("\t█████████");
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("█████████");
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.Write("█████████");

            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.Write("\t███████████████████████████");

            Console.Write("\n");

            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("\t█████████");
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("█████████");
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.Write("█████████");

            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.Write("\t███████████████████████████");

            Console.Write("\n\n\n");

            Console.ForegroundColor = ConsoleColor.White;       //Begrüssung
            Console.WriteLine("Herzlich Willkommen zu unserer Wörtliabfragesystem, indem man Deutsch und Französisch \nVokabeln üben kann.\n");
            Console.WriteLine("Um Französisch-Deutsch zu üben, geben Sie 'frde' ein.");
            Console.WriteLine("Um Deutsch-Französisch zu üben, geben Sie 'defr' ein.\n\n");

            string start;

            start = Console.ReadLine();   //eingabe um zu starten

            if (start == "frde")    //auf Französisch - Deutsch setzen
            {

                Console.Clear();
                franzDeutsch();
            }
            else if (start == "defr")   //auf Deutsch - Französisch setzen
            {
                Console.Clear();
                deutschFranz();
            }
            else         //Fehlerbehebung 
            {
                Console.WriteLine("Ungültiger Befehl! Geben Sie entweder 'frde' für Französisch-Deutsch oder 'defr' um Deutsch-Französisch.");
                Console.WriteLine("Drücken Sie 'Enter' um fortzufahren.");
                Console.ReadLine();
                Console.Clear();
                anfang();
            }
        }

        public static void deutschFranz()       //Deutsch - Französisch
        {
            Application excelApp = new Application();

            if (excelApp == null)           //Falls Excel nicht instaliert wäre.
            {
                Console.WriteLine("Excel ist nicht installiert.");
                return;
            }

            Workbook excelBook = excelApp.Workbooks.Open(@"Geben Sie hier das Pfad des Exceldateis an.");      //Referenz aufs Exceldatei und öffnen vom Exceldatei. Wichtig!: Pfad eingeben!           
            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;        //Einlesen von Exceldatei

            int rows = excelRange.Rows.Count;               //Zellenzähler
            int cols = excelRange.Columns.Count;

            maxPunktZahl = rows - 1;        //Maximal erreichbare Punktzahl gleich mit dem int rows, also mit dem Anzahl ausgefüllten Reihen ausser die obersten.

            string Lösung = excelRange.Cells[reihe, 2].Value2;          //Referenz auf Zelle mit dem Lösung
            string ausgabe = excelRange.Cells[reihe, 1].Value2;        //Referenz auf Zelle mit dem Ausgabe

            if (ausgabe == null)            //Falls die Zelle mit dem Ausgabe leer wäre.
            {
                Console.Clear();
                schluss();                  //Führt zum Schlusswort.
            }
            else if (ausgabe != null)       //Falls die Zelle mit dem Ausgabe nicht leer wäre.
            {
                Console.WriteLine($"Score: " + punktZahl + " Punkte von " + maxPunktZahl + "\n");    //Ausgabe Punktzahl        
                Console.Write(ausgabe + "\t");      //Ausgabe Zelleninhalt
                string eingabe;                     //Eingabe

                eingabe = Console.ReadLine();

                excelApp.Quit();                //Loslassen von Exceldatei nach dem einlesen um Ausnahmen und offenne Excelfenster zu vermeiden.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                if (eingabe != Lösung)          //Falls das eingegebene Wort falsch wäre.
                {
                    Console.WriteLine("Falsch!\n");
                    deutschFranz();             //Führt zu einer neuer Abfrage mit dem gleichem Wort.
                }
                else if (eingabe == Lösung)     //Falls das eingegebene Wort richtig wäre. 
                {
                    Console.WriteLine("Richtig!\n");
                    reihe++;                    //Führt zur nächsten Reihe im Exceldatei.
                    punktZahl++;                //Vergrössert die Punktzahl um eins.
                    deutschFranz();             //Führt zu einer neuer Abfrage mit dem nächstem Wort.
                }
            }
        }

        public static void franzDeutsch()   //Französisch - Deutsch
        {
            Application excelApp = new Application();

            if (excelApp == null)           //Falls Excel nicht instaliert wäre.
            {
                Console.WriteLine("Excel ist nicht installiert.");
                return;
            }

            Workbook excelBook = excelApp.Workbooks.Open(@"Geben Sie hier das Pfad des Exceldateisan.");      //Referenz aufs Exceldatei und öffnen vom Exceldatei. Wichtig!: Pfad eingeben!           
            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;        //Einlesen von Exceldatei

            int rows = excelRange.Rows.Count;               //Zellenzähler
            int cols = excelRange.Columns.Count;

            maxPunktZahl = rows - 1;        //Maximal erreichbare Punktzahl gleich mit dem int rows, also mit dem Anzahl ausgefüllten Reihen ausser die obersten.

            string Lösung = excelRange.Cells[reihe, 1].Value2;          //Referenz auf Zelle mit dem Lösung
            string ausgabe = excelRange.Cells[reihe, 2].Value2;        //Referenz auf Zelle mit dem Ausgabe

            if (ausgabe == null)            //Falls die Zelle mit dem Ausgabe leer wäre.
            {
                Console.Clear();
                schluss();                  //Führt zum Schlusswort.
            }
            else if (ausgabe != null)       //Falls die Zelle mit dem Ausgabe nicht leer wäre.
            {
                Console.WriteLine($"Score: " + punktZahl + " Punkte von " + maxPunktZahl + "\n");    //Ausgabe Punktzahl        
                Console.Write(ausgabe + "\t");      //Ausgabe Zelleninhalt
                string eingabe;                     //Eingabe

                eingabe = Console.ReadLine();

                excelApp.Quit();                //Loslassen von Exceldatei nach dem einlesen um Ausnahmen und offenne Excelfenster zu vermeiden.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                if (eingabe != Lösung)          //Falls das eingegebene Wort falsch wäre.
                {
                    Console.WriteLine("Falsch!\n");
                    franzDeutsch();             //Führt zu einer neuer Abfrage mit dem gleichem Wort.
                }
                else if (eingabe == Lösung)     //Falls das eingegebene Wort richtig wäre. 
                {
                    Console.WriteLine("Richtig!\n");
                    reihe++;                    //Führt zur nächsten Reihe im Exceldatei.
                    punktZahl++;                //Vergrössert die Punktzahl um eins.
                    franzDeutsch();             //Führt zu einer neuer Abfrage mit dem nächstem Wort.
                }
            }
        }

        public static void schluss()        //Schluss des Programmes mit dem Ausgabe der erreichten Punkten und maximalen Punkten. 
        {
            Console.WriteLine($"Gratuliere, du hast " + punktZahl + " Punkte von " + maxPunktZahl + " erreicht.\n");
            Console.WriteLine("Drücken Sie eine Beliebige Taste um die Kosole zu schliessen.");
            Console.ReadKey();
        }
    }
}
