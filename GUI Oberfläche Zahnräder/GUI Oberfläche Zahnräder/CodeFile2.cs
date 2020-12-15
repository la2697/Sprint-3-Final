using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;



namespace GUI_Oberfläche_Zahnräder
{
    public class Außenverzahnung
    {
        //Eingabevariablen
        public double m;                            //Modul
        public double c;                            //Kopfspiel Standard
        public double d;                            //Teilkreisdurchmesser
        public double z;                            //Zähnezahl
        public double t;                            //Werkstoffdicke
        

        //Help
        public int Fehler;

        //Werkstoffangaben
        public double Preis_S235JR = 3.63;          //Preis S235JR
        public double Preis_GJL = 0.80;             //Preis GJL
        public double Preis_GJS = 1.10;             //Preis GJS


        //Ausgabevariablen
        public double h;                            //Zahnhöhe
        public double ha;                           //Zahnkopfhöhe
        public double hf;                           //Zahnfußhöhe
        public double df;                           //Fußkreisdurchmesser
        public double da;                           //Kopfkreisdurchmesser
        public double p;                            //Teilung
        public double material;                     //Materialdichte g/cm3
        public double preis;                        //Preis pro kg
        public double A;                            //Flächeninhalt mm2
        public double V;                            //Volumen mm3
        public double M;                            //Masse g
        public double Wert;                         //Preis




        public void Berechnung()
        {
            //Formeln
            d = m * z;
            h = (2 * m) + c;
            p = Math.PI * m;
            da = d + (2 + m);
            ha = m;
            hf = m + c;
            df = d - 2 * (m + c);
            A = (Math.PI * d * d) / 4;
            V = A * t;
            M = V * 1 / 1000 * material;
            Wert = M * 1 / 1000 * preis;

            //Rundung
            d = Math.Round(d, 2);
            p = Math.Round(p, 2);
            h = Math.Round(h, 2);
            da = Math.Round(da, 2);
            ha = Math.Round(ha, 2);
            df = Math.Round(df, 2);
            A = Math.Round(A, 2);
            V = Math.Round(V, 2);
            M = Math.Round(M, 2);
            Wert = Math.Round(Wert, 2);
        }

    }
}