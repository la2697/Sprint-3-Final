using System;
using System.Windows;
using INFITF;
using MECMOD;
using PARTITF;
using HybridShapeTypeLib;
using KnowledgewareTypeLib;
using System.Windows.Shapes;
using ProductStructureTypeLib;


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
            double Beta = 140 / av.z;
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
            Point2D point_Ursprung = catfactory2D1.CreatePoint(x0, y0);                                                         
            Point2D point_KopfkreisLinks = catfactory2D1.CreatePoint(xKopfkreis, yKopfkreis);                                   
            Point2D point_FußkreisLinks = catfactory2D1.CreatePoint(xFußkreis, yFußkreis);                                    
            Point2D point_KopfkreisRechts = catfactory2D1.CreatePoint(-xKopfkreis, yKopfkreis);
            Point2D point_FußkreisRechts = catfactory2D1.CreatePoint(-xFußkreis, yFußkreis);
            Point2D point_AnfangspunktLinks = catfactory2D1.CreatePoint(-x_AnfangspunktFußkreis, y_AnfangspunktFußkreis);

            //Linien
            Line2D line_FußkreisKopfkreis = catfactory2D1.CreateLine(xFußkreis, yFußkreis, xKopfkreis, yKopfkreis);
            line_FußkreisKopfkreis.StartPoint = point_FußkreisLinks;
            line_FußkreisKopfkreis.EndPoint = point_KopfkreisLinks;

            Line2D line_KopfkreisFußkreis = catfactory2D1.CreateLine(-xKopfkreis, yKopfkreis, -xFußkreis, yFußkreis);
            line_KopfkreisFußkreis.StartPoint = point_KopfkreisRechts;
            line_KopfkreisFußkreis.EndPoint = point_FußkreisRechts;


            //Kreise
            Circle2D circle_KopfkreisLinksRechts = catfactory2D1.CreateCircle(x0, y0, Kopfkreisradius, 0, Math.PI * 2);
            circle_KopfkreisLinksRechts.CenterPoint = point_Ursprung;
            circle_KopfkreisLinksRechts.EndPoint = point_KopfkreisLinks;
            circle_KopfkreisLinksRechts.StartPoint = point_KopfkreisRechts;

            Circle2D circle_AnfangFußkreis = catfactory2D1.CreateCircle(x0, x0, Fußkreisradius, 0, Math.PI * 2);
            circle_AnfangFußkreis.CenterPoint = point_Ursprung;
            circle_AnfangFußkreis.EndPoint = point_AnfangspunktLinks;
            circle_AnfangFußkreis.StartPoint = point_FußkreisLinks;

            hsp_catiaProfil.CloseEdition();

            hsp_catiaPart.Part.Update();


            //Profilerstellen Ende

            //Kreismuster

            ShapeFactory SF = (ShapeFactory)hsp_catiaPart.Part.ShapeFactory;
            HybridShapeFactory HSF = (HybridShapeFactory)hsp_catiaPart.Part.HybridShapeFactory;
            Part myPart = hsp_catiaPart.Part;

            Factory2D Factory2D2 = hsp_catiaProfil.Factory2D;
            HybridShapePointCoord Ursprung = HSF.AddNewPointCoord(0, 0, 0);
            Reference RefUrsprung = myPart.CreateReferenceFromObject(Ursprung);
            HybridShapeDirection XDir = HSF.AddNewDirectionByCoord(1, 0, 0);
            Reference RefXDir = myPart.CreateReferenceFromObject(XDir);

            CircPattern Kreismuster = SF.AddNewSurfacicCircPattern(Factory2D2, 1, 2, 0, 0, 1, 1, RefUrsprung, RefXDir, false, 0, true, false);
            Kreismuster.CircularPatternParameters = CatCircularPatternParameters.catInstancesandAngularSpacing;
            AngularRepartition angularRepartition1 = Kreismuster.AngularRepartition;
            Angle angle1 = angularRepartition1.AngularSpacing;
            angle1.Value = Convert.ToDouble(360 / Convert.ToDouble(av.z));
            AngularRepartition angularRepartition2 = Kreismuster.AngularRepartition;
            IntParam intParam1 = angularRepartition2.InstancesCount;
            intParam1.Value = Convert.ToInt32(av.z) + 1;

            Reference Ref_Kreismuster = myPart.CreateReferenceFromObject(Kreismuster);
            HybridShapeAssemble Verbindung = HSF.AddNewJoin(Ref_Kreismuster, Ref_Kreismuster);
            Reference Ref_Verbindung = myPart.CreateReferenceFromObject(Verbindung);
            HSF.GSMVisibility(Ref_Verbindung, 0);
            myPart.Update();
            Bodies bodies = myPart.Bodies;
            Body myBody = bodies.Add();
            myBody.set_Name("Zahnrad");
            myBody.InsertHybridShape(Verbindung);
            myPart.Update();

            myPart.InWorkObject = myBody;
            Pad myPad = SF.AddNewPadFromRef(Ref_Verbindung, av.t);
            myPart.Update();


            //Bohrung
            Reference RefBohrung1 = hsp_catiaPart.Part.CreateReferenceFromBRepName("FSur:(Face:(Brp:(Pad.1;2);None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", myPad);
            Hole catBohrung1 = SF.AddNewHoleFromPoint(0, 0, 0, RefBohrung1, 0);
            Length catLengthBohrung1 = catBohrung1.Diameter;
            catLengthBohrung1.Value = Convert.ToDouble(av.d / 2);

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
