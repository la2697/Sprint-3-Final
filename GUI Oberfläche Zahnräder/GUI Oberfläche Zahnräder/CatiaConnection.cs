using System;
using System.Windows;
using INFITF;
using MECMOD;
using PARTITF;
using HybridShapeTypeLib;
using KnowledgewareTypeLib;


namespace GUI_Oberfläche_Zahnräder
{
    class CatiaConnection
    {
        INFITF.Application hsp_catiaApp;
        MECMOD.PartDocument hsp_catiaPart;
        MECMOD.Sketch hsp_catiaProfil;
        
        public bool CATIALaeuft()
        {
            try
            {
                object catiaObject = System.Runtime.InteropServices.Marshal.GetActiveObject(
                    "CATIA.Application");
                hsp_catiaApp = (INFITF.Application) catiaObject;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        //Test
        public bool ErzeugePart()
        {
            INFITF.Documents catDocuments1 = hsp_catiaApp.Documents;
            hsp_catiaPart = catDocuments1.Add("Part") as MECMOD.PartDocument;
            return true;
        }

        private void ErzeugeAchsensystem()
        {
            object[] arr = new object[] {0.0, 0.0, 0.0,
                                         0.0, 1.0, 0.0,
                                         0.0, 0.0, 1.0 };
            hsp_catiaProfil.SetAbsoluteAxisData(arr);
        }

        public void ErstelleLeereSkizze()
        {
            // geometrisches Set auswaehlen und umbenennen
            HybridBodies catHybridBodies1 = hsp_catiaPart.Part.HybridBodies;
            HybridBody catHybridBody1;
            try
            {
                catHybridBody1 = catHybridBodies1.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                MessageBox.Show("Kein geometrisches Set gefunden! " + Environment.NewLine +
                    "Ein PART manuell erzeugen und ein darauf achten, dass 'Geometisches Set' aktiviert ist.",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            catHybridBody1.set_Name("Profile");
            // neue Skizze im ausgewaehlten geometrischen Set anlegen
            Sketches catSketches1 = catHybridBody1.HybridSketches;
            OriginElements catOriginElements = hsp_catiaPart.Part.OriginElements;
            Reference catReference1 = (Reference)catOriginElements.PlaneYZ;
            hsp_catiaProfil = catSketches1.Add(catReference1);

            // Achsensystem in Skizze erstellen 
            ErzeugeAchsensystem();

            // Part aktualisieren
            hsp_catiaPart.Part.Update();
        }

        public void Stirnzahnrad(MainWindow.Außenverzahnung av)
        {
            
            

            //Profil Erstellen

            //Nullpunkt
            double x0 = 0;
            double y0 = 0;

            //Hilfsgrößen
            double Teilkreisradius = av.d / 2;
            double Hilfskreisradius = Teilkreisradius * 0.94;
            double Fußkreisradius = Teilkreisradius - (1.25 * av.m);
            double Kopfkreisradius = Teilkreisradius + av.m;
            double Verrundungsradius = 0.35 * av.m;

            double Alpha = 2;
            double Alpharad = Math.PI * Alpha / 180;
            double Beta = 120 / av.z;
            double Betarad = Math.PI * Beta / 180;
            double Gamma = 90 - (Alpha - Beta);
            double Gammarad = Math.PI * Gamma / 180;
            double Totalangel = 360.0 / av.z;
            double Totalangelrad = Math.PI * Totalangel / 180;


            //Punkte
            //Kopfkreis
            double xKopfkreis = -Kopfkreisradius * Math.Sin(Alpharad);
            double yKopfkreis = Kopfkreisradius * Math.Cos(Alpharad);
  
            //Fußkreis
            double xFußkreis = -Fußkreisradius * Math.Sin(Betarad);
            double yFußkreis = Fußkreisradius * Math.Cos(Betarad);

            //Koordinaten Anfangspunkt Fußkreis
            double Hilfswinkel = Totalangelrad - Math.Atan(Math.Abs(xFußkreis) / Math.Abs(yFußkreis));
            double x_AnfangspunktFußkreis = Fußkreisradius * Math.Sin(Hilfswinkel);
            double y_AnfangspunktFußkreis = Fußkreisradius * Math.Cos(Hilfswinkel);

            //Skizze umbenennen und öffnen 
            hsp_catiaProfil.set_Name("Zahnradskizze");
            Factory2D catfactory2D1 = hsp_catiaProfil.OpenEdition();

            //Nun die Punkte in die Skizze
            Point2D point_Ursprung = catfactory2D1.CreatePoint(x0, y0);                                                         //PKT1
            Point2D point_KopfkreisLinks = catfactory2D1.CreatePoint(xKopfkreis, yKopfkreis);
            Point2D point_FußkreisLinks = catfactory2D1.CreatePoint(xFußkreis, yFußkreis);
            Point2D point_KopfkreisRechts = catfactory2D1.CreatePoint(-xKopfkreis, yKopfkreis);
            Point2D point_FußkreisRechts = catfactory2D1.CreatePoint(-xFußkreis, yFußkreis);
            Point2D point_AnfangspunktLinks = catfactory2D1.CreatePoint(-x_AnfangspunktFußkreis, y_AnfangspunktFußkreis);

            //Erzeuge Linien
            Line2D line_Kopfkreis = catfactory2D1.CreateLine(xKopfkreis, yKopfkreis, -xKopfkreis, yKopfkreis);
            line_Kopfkreis.StartPoint = point_KopfkreisLinks;
            line_Kopfkreis.EndPoint = point_KopfkreisRechts;

            Line2D line_FußkreisKopfkreisLinks = catfactory2D1.CreateLine(xFußkreis, yFußkreis, xKopfkreis, yKopfkreis);
            line_FußkreisKopfkreisLinks.StartPoint = point_FußkreisLinks;
            line_FußkreisKopfkreisLinks.EndPoint = point_KopfkreisLinks;

            Line2D line_FußkreisKopfkreisRechts = catfactory2D1.CreateLine(xFußkreis, yFußkreis, xKopfkreis, yKopfkreis);
            line_FußkreisKopfkreisRechts.StartPoint = point_FußkreisRechts;
            line_FußkreisKopfkreisRechts.EndPoint = point_KopfkreisRechts;

            //Kreise
            Circle2D circle_AnfangFußkreis = catfactory2D1.CreateCircle(x0, y0, Fußkreisradius, -x_AnfangspunktFußkreis, y_AnfangspunktFußkreis);
            circle_AnfangFußkreis.CenterPoint = point_Ursprung;
            circle_AnfangFußkreis.EndPoint = point_AnfangspunktLinks;
            circle_AnfangFußkreis.StartPoint = point_FußkreisLinks;

            
            hsp_catiaProfil.CloseEdition();

            hsp_catiaPart.Part.Update();

            Factory2D catfactory2D2 = hsp_catiaProfil.OpenEdition();
            hsp_catiaPart.Part.Update();
            hsp_catiaProfil.CloseEdition();

            //Profilerstellen Ende

            //Kreistmuster

            ShapeFactory shapeFactory1 = (ShapeFactory)hsp_catiaPart.Part.ShapeFactory;
                HybridShapeFactory hybridShapeFactory1 = (HybridShapeFactory)hsp_catiaPart.Part.HybridShapeFactory;

                Factory2D factory2D2 = hsp_catiaProfil.Factory2D;

                HybridShapePointCoord ursprung = hybridShapeFactory1.AddNewPointCoord(0, 0, 0);
                Reference refUrsprung = hsp_catiaPart.Part.CreateReferenceFromObject(ursprung);

                HybridShapeDirection xRichtung = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0);
                Reference refxRichtung = hsp_catiaPart.Part.CreateReferenceFromObject(xRichtung);

                CircPattern kreismuster = shapeFactory1.AddNewSurfacicCircPattern(factory2D2, 1, 2, 0, 0, 1, 1, refUrsprung, refxRichtung, false, 0, true, false);
                kreismuster.CircularPatternParameters = CatCircularPatternParameters.catInstancesandAngularSpacing;
                AngularRepartition angularRepartition1 = kreismuster.AngularRepartition;
                Angle angle1 = angularRepartition1.AngularSpacing;
                angle1.Value = Convert.ToDouble(360 / av.z);
                AngularRepartition angularRepartition2 = kreismuster.AngularRepartition;
                IntParam intParam1 = angularRepartition2.InstancesCount;
                intParam1.Value = Convert.ToInt32(av.z)+1;

                //Kreismusterenden verbinden

                Reference refKreismuster = hsp_catiaPart.Part.CreateReferenceFromObject(kreismuster);
                HybridShapeAssemble verbindung = hybridShapeFactory1.AddNewJoin(refKreismuster, refKreismuster);
                Reference refVerbindung = hsp_catiaPart.Part.CreateReferenceFromObject(verbindung);

                hybridShapeFactory1.GSMVisibility(refVerbindung, 0);

                hsp_catiaPart.Part.MainBody.InsertHybridShape(verbindung);


            //Part aktualisieren
                hsp_catiaPart.Part.Update();

                

        }

        public void Modellerzeugung()
        {
            // Hauptkoerper in Bearbeitung definieren
            hsp_catiaPart.Part.InWorkObject = hsp_catiaPart.Part.MainBody;

            // Block erzeugen
            ShapeFactory catShapeFactory1 = (ShapeFactory)hsp_catiaPart.Part.ShapeFactory;
            Pad catPad1 = catShapeFactory1.AddNewPad(hsp_catiaProfil, 2);

            // Block umbenennen
            catPad1.set_Name("Zahnrad-Modell");

            // Part aktualisieren
            hsp_catiaPart.Part.Update();
        }

        private double Schnittpunkt_X(double xMittelpunkt, double yMittelpunkt, double Radius1, double xMittelpunkt2, double yMittelpunkt2, double Radius2)
        {
            double d = Math.Sqrt(Math.Pow((xMittelpunkt - xMittelpunkt2), 2) + Math.Pow((yMittelpunkt - yMittelpunkt2), 2));
            double l = (Math.Pow(Radius1, 2) - Math.Pow(d, 2)) / (d * 2);
            double h;
            double Verbindungsabfrage = 0.00001;
            
            if (Radius1 - 1 < -Verbindungsabfrage)
            {
                MessageBox.Show("Fehler Verbindungsabfrage");
            }
            if (Math.Abs(Radius1 - 1) < Verbindungsabfrage)
            {
                h = 0;
            }
            else
            {
                h = Math.Sqrt(Math.Pow(Radius1, 2) - Math.Pow(1, 2));
            }

            return 1 * (xMittelpunkt2 - xMittelpunkt) / d - h * (yMittelpunkt2 - yMittelpunkt) / d + xMittelpunkt;
        }


        private double Schnittpunkt_Y(double xMittelpunkt, double yMittelpunkt, double Radius1, double xMittelpunkt2, double yMittelpunkt2, double Radius2)
        {
            double d = Math.Sqrt(Math.Pow((xMittelpunkt - xMittelpunkt2), 2) + Math.Pow((yMittelpunkt - yMittelpunkt2), 2));
            double l = (Math.Pow(Radius1, 2) - Math.Pow(Radius2, 2) + Math.Pow(d, 2)) / (d * 2);
            double h;
            double Verbindungsabfrage = 0.00001;

            if (Radius1 - 1 < -Verbindungsabfrage)
            {
                MessageBox.Show("Fehler Verbindungsabfrage 2");
            }
            if (Math.Abs(Radius1 - 1) < Verbindungsabfrage)
            {
                h = 0;
            }
            else
            {
                h = Math.Sqrt(Math.Pow(Radius1, 2) - Math.Pow(1, 2));
            }

            return 1 * (yMittelpunkt2 - yMittelpunkt) / d - h * (xMittelpunkt2 - xMittelpunkt) / d + yMittelpunkt;
        }

    }
}
