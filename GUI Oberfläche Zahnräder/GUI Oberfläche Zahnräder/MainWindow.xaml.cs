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
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>



    public partial class MainWindow : Window
    {

        //Zahlencheck
        public bool Zahlenprüfung(string Zahlencheck)
        {
            try
            {
                double doublezahl = double.Parse(Zahlencheck);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

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

        public class Innenverzahnung
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
            public double dm;                           //Mittlerer Durchmesser

            public void Berechnung()
            {
                //Formeln
                d = m * z;
                h = (2 * m) + c;
                p = Math.PI * m;
                da = d - (2 * m);
                ha = m;
                hf = m + c;
                df = d + 2 * (m + c);
                dm = d * 1.4;
                A = ((Math.PI * ((dm * dm)-(da * da)) / 4) - (Math.PI * m * h * z) / 2);
                V = A * t;
                M = V * 1 / 1000 * material;
                Wert = M * 1 / 1000 * preis;

                //Rundung
                d = Math.Round(d, 2);
                dm = Math.Round(dm, 2);
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

        public class Schrägverzahnung
        {
            //Eingabevariablen
            public double mt;                           //Stirnmodul
            public double mn;                           //Normalmodul
            public double pt;                           //Stirnteilung
            public double pn;                           //Normalteilung
            public double c;                            //Kopfspiel Standard
            public double d;                            //Teilkreisdurchmesser
            public double z;                            //Zähnezahl
            public double t;                            //Werkstoffdicke
            public double beta;                         //Schrägungswinkel (8°-25°)


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
            public double dm;                           //Mittlerer Durchmesser

            public void Berechnung()
            {
                //Formeln
                mt = mn / Math.Cos(beta);
                d = mt * z;
                h = (2 * mn) + c;
                pn = Math.PI * mn;
                pt = pn / Math.Cos(beta);
                da = d + (2 * mn);
                ha = mn;
                hf = mn + c;
                df = d + 2 * (mn + c);
                dm = d * 1.4;
                A = ((Math.PI * ((dm * dm) - (da * da)) / 4) - (Math.PI * mn * h * z) / 2);
                V = A * t;
                M = V * 1 / 1000 * material;
                Wert = M * 1 / 1000 * preis;

                //Rundung
                d = Math.Round(d, 2);
                mt = Math.Round(mt, 2);
                pt = Math.Round(pt, 2);
                dm = Math.Round(dm, 2);
                pn = Math.Round(p, 2);
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

        public MainWindow()
        {

        }

        public void btn_Bestätigen_Click(object sender, RoutedEventArgs e)
        {
            //Außenverzahnung
           
            if (rtb_Gerad.IsChecked == true)
            {
                
                Außenverzahnung av = new Außenverzahnung();  
                switch (CB_Werkstoff.Text)
                {
                    case "GJL":
                        av.material = 7.15;
                        av.preis = 0.80;
                        break;

                    case "GJS":
                        av.material = 7.1;
                        av.preis = 1.10;
                        break;

                    case "S235JR":
                        av.material = 7.85;
                        av.preis = 3.63;
                        break;
                }

                //Eingabeparameter Anzeige
                lb_Werkstoffwahl.Content = CB_Werkstoff.Text;
                lb_Zähne.Content = txt_Zähnezahl.Text;
                lb_Kopfspiele1.Content = TB_0167.Text;
                lb_Modul2.Content = txt_Modul.Text;
                lb_Dicke1.Content = txt_Dicke.Text;
                
                

             
                //Eingabechecks
                //Eingabecheck Modul
                string Zahlencheck = txt_Modul.Text;
                if (Zahlenprüfung(Zahlencheck) == true)
                {
                    av.m = Convert.ToDouble(txt_Modul.Text);
                    if (av.m <= 0)
                    {
                        MessageBox.Show("Fehler! Modul kann nur eine positive Zahl betragen.");
                        av.Fehler = 1;
                    }
                }
                else
                {
                    MessageBox.Show("Fehler! Modul kann nur eine positive Zahl betragen.");
                    av.Fehler = 1;
                }

                //Kopfspiel
                Zahlencheck = TB_0167.Text;
                if (Zahlenprüfung(Zahlencheck) == true)
                {
                    av.c = Convert.ToDouble(TB_0167.Text);
                    if ((av.c < 0.1) || (av.c > 0.3))
                    {
                        MessageBox.Show("Fehler! Kopfspiel kann nur eine Zahl zwischen 0,1 und 0,3 betragen.");
                        av.Fehler = 1;
                    }
                }
                else
                {
                    MessageBox.Show("Fehler! Kopfspiel kann nur eine Zahl zwischen 0,1 und 0,3 betragen.");
                    av.Fehler = 1;
                }


                //Eingabecheck Zähnezahl
                Zahlencheck = txt_Zähnezahl.Text;
                if (Zahlenprüfung(Zahlencheck) == true)
                {
                    av.z = Convert.ToDouble(txt_Zähnezahl.Text);
                    if (av.z <= 0)
                    {
                        MessageBox.Show("Fehler! Zähnezahl kann nur eine positive Ganzzahl ab 5 betragen.");
                        av.Fehler = 1;
                    }
                }
                else
                {
                    MessageBox.Show("Fehler! Zähnezahl kann nur eine positive Ganzzahl ab 5 betragen.");
                    av.Fehler = 1;
                }

                //Eingabecheck Dicke
                Zahlencheck = txt_Dicke.Text;
                if (Zahlenprüfung(Zahlencheck) == true)
                {
                    av.t = Convert.ToDouble(txt_Dicke.Text);
                    if (av.t <= 0)
                    {
                        MessageBox.Show("Fehler! Dicke kann nur eine positive Zahl betragen.");
                        av.Fehler = 1;
                    }
                }
                else
                {
                    MessageBox.Show("Fehler! Dicke kann nur eine positive Zahl betragen.");
                    av.Fehler = 1;
                }

                //Berechnung durchführen
                if (av.Fehler == 0)
                {
                    av.Berechnung();
                    Lbl_Zahnhöhe.Content = Convert.ToString(av.h);
                    lbl_Fußkreisdurchmesser.Content = Convert.ToString(av.df);
                    lbl_Teilkreisdurchmesser.Content = Convert.ToString(av.d);
                    lbl_Zahnfußhöhe.Content = Convert.ToString(av.hf);
                    lbl_Zahnkopfhöhe.Content = Convert.ToString(av.ha);
                    lbl_Teilung.Content = Convert.ToString(av.p);
                    lbl_Kopfkreisdurchmesser.Content = Convert.ToString(av.da);
                    lbl_Gewicht.Content = Convert.ToString(av.M);
                    lbl_Preis.Content = Convert.ToString(av.Wert);
                    MessageBox.Show(Convert.ToString(av.c));

                    //Felder ausbleichen
                    lbl_stirnmodul.IsEnabled = false;
                    lbl_stirnteilung.IsEnabled = false;
                    lb_Stirnmodul.IsEnabled = false;
                    lb_Stirnteilung.IsEnabled = false;
                    lb_Schrägungswinkel.IsEnabled = false;
                    lb_Winkel.IsEnabled = false;

                }
            }

            //Innenverzahung
            else if(rtb_Innenzahnrad.IsChecked == true)
            {
                
                Innenverzahnung iv = new Innenverzahnung();
                switch (CB_Werkstoff.Text)
                {
                    case "GJL":
                        iv.material = 7.15;
                        iv.preis = 0.80;
                        break;

                    case "GJS":
                        iv.material = 7.1;
                        iv.preis = 1.10;
                        break;

                    case "S235JR":
                        iv.material = 7.85;
                        iv.preis = 3.63;
                        break;
                }

                //Eingabeparameter Anzeige
                lb_Werkstoffwahl.Content = CB_Werkstoff.Text;
                lb_Zähne.Content = txt_Zähnezahl.Text;
                lb_Kopfspiele1.Content = TB_0167.Text;
                lb_Modul2.Content = txt_Modul.Text;
                lb_Dicke1.Content = txt_Dicke.Text;

                //Eingabechecks
                //Modul
                string Zahlencheck = txt_Modul.Text;
                if (Zahlenprüfung(Zahlencheck) == true)
                {
                    iv.m = Convert.ToDouble(txt_Modul.Text);
                    if (iv.m <= 0)
                    {
                        MessageBox.Show("Fehler! Modul kann nur eine positive Zahl betragen.");
                        iv.Fehler = 1;
                    }
                }
                else
                {
                    MessageBox.Show("Fehler! Modul kann nur eine positive Zahl betragen.");
                    iv.Fehler = 1;
                }

                //Zähnezahl
                Zahlencheck = txt_Zähnezahl.Text;
                if (Zahlenprüfung(Zahlencheck) == true)
                {
                    iv.z = Convert.ToDouble(txt_Zähnezahl.Text);
                    if (iv.z <= 0)
                    {
                        MessageBox.Show("Fehler! Zähnezahl kann nur eine positive Ganzzahl ab 5 betragen.");
                        iv.Fehler = 1;
                    }
                }
                else
                {
                    MessageBox.Show("Fehler! Zähnezahl kann nur eine positive Ganzzahl ab 5 betragen.");
                    iv.Fehler = 1;
                }

                //Eingabecheck Dicke
                Zahlencheck = txt_Dicke.Text;
                if (Zahlenprüfung(Zahlencheck) == true)
                {
                    iv.t = Convert.ToDouble(txt_Dicke.Text);
                    if (iv.t <= 0)
                    {
                        MessageBox.Show("Fehler! Dicke kann nur eine positive Zahl betragen.");
                        iv.Fehler = 1;
                    }
                }
                else
                {
                    MessageBox.Show("Fehler! Dicke kann nur eine positive Zahl betragen.");
                    iv.Fehler = 1;
                }

                //Kopfspiel
                Zahlencheck = TB_0167.Text;
                if (Zahlenprüfung(Zahlencheck) == true)
                {
                    iv.c = Convert.ToDouble(TB_0167.Text);
                    if ((iv.c < 0.1) || (iv.c > 0.3))
                    {
                        MessageBox.Show("Fehler! Kopfspiel kann nur eine Zahl zwischen 0,1 und 0,3 betragen.");
                        iv.Fehler = 1;
                    }
                }
                else
                {
                    MessageBox.Show("Fehler! Kopfspiel kann nur eine Zahl zwischen 0,1 und 0,3 betragen.");
                    iv.Fehler = 1;
                }

                if (iv.Fehler == 0)
                {
                    iv.Berechnung();
                    Lbl_Zahnhöhe.Content = Convert.ToString(iv.h);
                    lbl_Fußkreisdurchmesser.Content = Convert.ToString(iv.df);
                    lbl_Teilkreisdurchmesser.Content = Convert.ToString(iv.d);
                    lbl_Zahnfußhöhe.Content = Convert.ToString(iv.hf);
                    lbl_Zahnkopfhöhe.Content = Convert.ToString(iv.ha);
                    lbl_Teilung.Content = Convert.ToString(iv.p);
                    lbl_Kopfkreisdurchmesser.Content = Convert.ToString(iv.da);
                    lbl_Gewicht.Content = Convert.ToString(iv.M);
                    lbl_Preis.Content = Convert.ToString(iv.Wert);

                    //Felder ausbleichen
                    lbl_stirnmodul.IsEnabled = false;
                    lbl_stirnteilung.IsEnabled = false;
                    lb_Stirnmodul.IsEnabled = false;
                    lb_Stirnteilung.IsEnabled = false;
                    lb_Schrägungswinkel.IsEnabled = false;
                    lb_Winkel.IsEnabled = false;
                }
            }

            //Schrägverzahnung
            else
            {
                Schrägverzahnung sv = new Schrägverzahnung();
                switch (CB_Werkstoff.Text)
                {
                    case "GJL":
                        sv.material = 7.15;
                        sv.preis = 0.80;
                        break;

                    case "GJS":
                        sv.material = 7.1;
                        sv.preis = 1.10;
                        break;

                    case "S235JR":
                        sv.material = 7.85;
                        sv.preis = 3.63;
                        break;
                }

                //Eingabeparameter Anzeige
                lb_Werkstoffwahl.Content = CB_Werkstoff.Text;
                lb_Zähne.Content = txt_Zähnezahl.Text;
                lb_Kopfspiele1.Content = TB_0167.Text;
                lb_Modul2.Content = txt_Modul.Text;
                lb_Dicke1.Content = txt_Dicke.Text;
                lb_Schrägungswinkel.Content = tb_Winkel.Text;

                //Eingabechecks
                //Modul
                string Zahlencheck = txt_Modul.Text;
                if (Zahlenprüfung(Zahlencheck) == true)
                {
                    sv.mn = Convert.ToDouble(txt_Modul.Text);
                    if (sv.mn <= 0)
                    {
                        MessageBox.Show("Fehler! Modul kann nur eine positive Zahl betragen.");
                        sv.Fehler = 1;
                    }
                }
                else
                {
                    MessageBox.Show("Fehler! Modul kann nur eine positive Zahl betragen.");
                    sv.Fehler = 1;
                }

                //Zähnezahl
                Zahlencheck = txt_Zähnezahl.Text;
                if (Zahlenprüfung(Zahlencheck) == true)
                {
                    sv.z = Convert.ToDouble(txt_Zähnezahl.Text);
                    if (sv.z <= 0)
                    {
                        MessageBox.Show("Fehler! Zähnezahl kann nur eine positive Ganzzahl ab 5 betragen.");
                        sv.Fehler = 1;
                    }
                }
                else
                {
                    MessageBox.Show("Fehler! Zähnezahl kann nur eine positive Ganzzahl ab 5 betragen.");
                    sv.Fehler = 1;
                }

                //Eingabecheck Dicke
                Zahlencheck = txt_Dicke.Text;
                if (Zahlenprüfung(Zahlencheck) == true)
                {
                    sv.t = Convert.ToDouble(txt_Dicke.Text);
                    if (sv.t <= 0)
                    {
                        MessageBox.Show("Fehler! Dicke kann nur eine positive Zahl betragen.");
                        sv.Fehler = 1;
                    }
                }
                else
                {
                    MessageBox.Show("Fehler! Dicke kann nur eine positive Zahl betragen.");
                    sv.Fehler = 1;
                }

                //Kopfspiel
                Zahlencheck = TB_0167.Text;
                if (Zahlenprüfung(Zahlencheck) == true)
                {
                    sv.c = Convert.ToDouble(TB_0167.Text);
                    if ((sv.c < 0.1) || (sv.c > 0.3))
                    {
                        MessageBox.Show("Fehler! Kopfspiel kann nur eine Zahl zwischen 0,1 und 0,3 betragen.");
                        sv.Fehler = 1;
                    }
                }
                else
                {
                    MessageBox.Show("Fehler! Kopfspiel kann nur eine Zahl zwischen 0,1 und 0,3 betragen.");
                    sv.Fehler = 1;
                }

                if (sv.Fehler == 0)
                {
                    sv.Berechnung();
                    Lbl_Zahnhöhe.Content = Convert.ToString(sv.h);
                    lbl_Fußkreisdurchmesser.Content = Convert.ToString(sv.df);
                    lbl_Teilkreisdurchmesser.Content = Convert.ToString(sv.d);
                    lbl_Zahnfußhöhe.Content = Convert.ToString(sv.hf);
                    lbl_Zahnkopfhöhe.Content = Convert.ToString(sv.ha);
                    lbl_Teilung.Content = Convert.ToString(sv.p);
                    lbl_Kopfkreisdurchmesser.Content = Convert.ToString(sv.da);
                    lbl_Gewicht.Content = Convert.ToString(sv.M);
                    lbl_Preis.Content = Convert.ToString(sv.Wert);
                    tb_Winkel.Text = Convert.ToString(sv.beta);
                }
            }

        }
    

        

        
        
            

        private void btn_Zurück_Click(object sender, RoutedEventArgs e)
        {
            //Felder leeren
            Lbl_Zahnhöhe.Content = " ";
            lbl_Fußkreisdurchmesser.Content = " ";
            lbl_Teilkreisdurchmesser.Content = " ";
            lbl_Zahnfußhöhe.Content = " ";
            lbl_Zahnkopfhöhe.Content = " ";
            lbl_Teilung.Content = " ";
            lbl_Kopfkreisdurchmesser.Content = " ";
            lbl_Gewicht.Content = " ";
            lbl_Preis.Content = " ";
            lb_Zähne.Content = " ";
            lb_Modul2.Content = " ";
            lb_Kopfspiele1.Content = " ";
            lb_Werkstoffwahl.Content = " ";
            lb_Dicke1.Content = " ";
            txt_Modul.Text = " ";
            txt_Zähnezahl.Text = " ";
            txt_Dicke.Text = " ";
            tb_Winkel.Text = " ";
            TB_0167.Text = "0,167";

            //ComboButton und RadioButton zurücksetzen
            S235JR.IsSelected = true;
            rtb_Gerad.IsChecked = true;         
            TB_0167.IsEnabled = false;
            CB_Kopfspiel.IsChecked = false;
            lbl_stirnmodul.IsEnabled = true;
            lbl_stirnteilung.IsEnabled = true;
            lb_Stirnmodul.IsEnabled = true;
            lb_Stirnteilung.IsEnabled = true;
            lb_Winkel.IsEnabled = true;
            lb_Schrägungswinkel.IsEnabled = true;
            
            //Silder zurücksetzen


        }


        //Info Zahnradparameter
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            GUI_Oberfläche_Zahnräder.Window1 Infofenster = new Window1();
            Infofenster.Show();
        }

        //Programm schließen
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Close();
        }

        //Catia Part erstellen
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            CatiaControl();
        }

        // // // // // // // // // // // // // // // //
        
        //CatiaControl
        
        
            private void CatiaControl()
            {
                try
                {

                    CatiaConnection cc = new CatiaConnection();

                    // Finde Catia Prozess
                    if (cc.CATIALaeuft())
                    {
                        Console.WriteLine("0");

                        // Öffne ein neues Part
                        cc.ErzeugePart();
                      

                        // Erstelle eine Skizze
                        cc.ErstelleLeereSkizze();
                        

                        // Generiere ein Profil
                        cc.Stirnzahnrad(av);
                        
                        
                       
                        
                    }
                    else
                    {
                        Console.WriteLine("Laufende Catia Application nicht gefunden");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Exception aufgetreten");
                }
                Console.WriteLine("Fertig - Taste drücken.");
                
            }

        private void CB_Kopfspiel_Checked(object sender, RoutedEventArgs e)
        {
            TB_0167.IsEnabled = true;
            TB_0167.Clear();
        }

        private void CB_Kopfspiel_Unchecked(object sender, RoutedEventArgs e)
        {
            TB_0167.IsEnabled = false;
            TB_0167.Text = "0,167";
        }

        private void rtb_Gerad_Click(object sender, RoutedEventArgs e)
        {
            tb_Winkel.IsEnabled = false;
            sliderbeta.IsEnabled = false;
            lb_Winkeltext.IsEnabled = false;
        }

        private void rtb_Schräg_Click(object sender, RoutedEventArgs e)
        {
            tb_Winkel.IsEnabled = true;
            sliderbeta.IsEnabled = true;
            lb_Winkeltext.IsEnabled = true;
        }

        private void rtb_Innenzahnrad_Click(object sender, RoutedEventArgs e)
        {
            tb_Winkel.IsEnabled = false;
            sliderbeta.IsEnabled = false;
            lb_Winkeltext.IsEnabled = false;
        }
    }


}


    




      


   




