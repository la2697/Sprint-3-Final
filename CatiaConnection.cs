using ProductStructureTypeLib;
using HybridShapeTypeLib;
using INFITF;
using KnowledgewareTypeLib;
using MECMOD;
using PARTITF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace WPF_Sprint2
{
    class CatiaConnection
    {
        INFITF.Application hsp_catiaApp;
        PartDocument hsp_catiaPart;
        Sketch hsp_catiaProfil;

        public bool CatiaLaeuft()
        {
            try
            {
                object catiaObj = System.Runtime.InteropServices.Marshal.GetActiveObject("CATIA.Application");
                hsp_catiaApp = (INFITF.Application)catiaObj;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public Boolean ErzeugePart()
        {
            Documents catDocuments = hsp_catiaApp.Documents;
            hsp_catiaPart = (PartDocument)catDocuments.Add("Part") as MECMOD.PartDocument;
            return true;
        }
        public void ErstelleLeereSkizze()
        {
            //geometrisches set auswählen und umbenennen
            HybridBodies catHybridBodies1 = hsp_catiaPart.Part.HybridBodies;
            HybridBody catHybridBody1;

            try
            {
                catHybridBody1 = catHybridBodies1.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                MessageBox.Show("Kein geometrisches Set gefunden!\nEin PART manuell erzeugen und darauf achten, dass ein 'Geometisches Set' aktiviert ist.", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            catHybridBody1.set_Name("Profile");

            //Neue Skizze im ausgewählten geometrischen Set anlegen
            Sketches catSketches1 = catHybridBody1.HybridSketches;
            OriginElements catOriginElements = hsp_catiaPart.Part.OriginElements;
            Reference catReference1 = (Reference)catOriginElements.PlaneYZ;
            hsp_catiaProfil = catSketches1.Add(catReference1);

            //Achsensystem in Skizze erzeugen
            ErzeugeAchsensystem();

            //Part aktualisieren
            hsp_catiaPart.Part.Update();
        }


        private void ErzeugeAchsensystem()
        {
            object[] arr = new object[] {0.0, 0.0, 0.0,
                                         0.0, 1.0, 0.0,
                                         0.0, 0.0, 1.0 };
            hsp_catiaProfil.SetAbsoluteAxisData(arr);
        }


        public void ErstelleProfilAußenverzahnung(Data dat)
        { 
            //HilfsRadien
            double d_r = dat.getTeilkreisdurchmesserZahnrad1() / 2;
            double hk_r = d_r * 0.94;
            double df_r = d_r - (1.25 * dat.getModulZahnrad1());
            double da_r = d_r + dat.getModulZahnrad1();
            double vd_r = 0.35 * dat.getModulZahnrad1();
            
            //HilfsWinkel
            double alpha = 20;
            double beta = 90 / dat.getZaehnezahlZahnrad1();
            double beta_r = Math.PI * beta / 180;
            double gamma = 90 - (alpha - beta);
            double gamma_r = Math.PI * gamma / 180;
            double ta = 360.0 / dat.getZaehnezahlZahnrad1();
            double ta_r = Math.PI * ta / 180;
            
            //Nullpunkte
            double x0 = 0;
            double y0 = 0;

            //MittelPunkt EvolventenKreis
            double MP_EvolventenKreis_x = hk_r * Math.Cos(gamma_r);
            double MP_EvolventenKreis_y = hk_r * Math.Sin(gamma_r);

            // SchnittPunkt Evolventenkreis & Teilkreisradius
            double SP_EvolventenTeilKreis_x = -d_r * Math.Sin(beta_r);
            double SP_EvolventenTeilKreis_y = d_r * Math.Cos(beta_r);

            //Evolventenkreis Radius
            double Evolventenkreis_r = Math.Sqrt(Math.Pow((MP_EvolventenKreis_x - SP_EvolventenTeilKreis_x), 2) + Math.Pow((MP_EvolventenKreis_y - SP_EvolventenTeilKreis_y), 2));

            //SchnittPunkt Evolventenkreis & Kopfkreisradius
            double SP_EvolventenKopfKreis_x = Schnittpunkt_x(x0, y0, da_r, MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r);
            double SP_EvolventenKopfKreis_y = Schnittpunkt_y(x0, y0, da_r, MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r);

            //MittelPunkt VerrundungsRadius
            double MP_Verrundung_x = Schnittpunkt_x(x0, y0, df_r + vd_r, MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r + vd_r);
            double MP_Verrundung_y = Schnittpunkt_y(x0, y0, df_r + vd_r, MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r + vd_r);

            //SchnittPunkt Evolventenkreis & Verrundungsradius
            double SP_EvolventeVerrundung_x = Schnittpunkt_x(MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r, MP_Verrundung_x, MP_Verrundung_y, vd_r);
            double SP_EvolventeVerrundung_y = Schnittpunkt_y(MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r, MP_Verrundung_x, MP_Verrundung_y, vd_r);

            //SchnittPunkt Fußkreis & Verrundungs Radius
            double SP_FußkreisVerrundungsRadius_x = Schnittpunkt_x(x0, y0, df_r, MP_Verrundung_x, MP_Verrundung_y, vd_r);
            double SP_FußkreisVerrundungsRadius_y = Schnittpunkt_y(x0, y0, df_r, MP_Verrundung_x, MP_Verrundung_y, vd_r);

            //StartPunkt Fußkreis Radius
            double phi = ta_r - Math.Atan(Math.Abs(SP_FußkreisVerrundungsRadius_x) / Math.Abs(SP_FußkreisVerrundungsRadius_y));
            double StartPkt_Fußkreis_x = -df_r * Math.Sin(phi);
            double StartPkt_Fußkreis_y = df_r * Math.Cos(phi);

            //Skizze umbenennen und öffnen
            hsp_catiaProfil.set_Name("AußenverzahnungEinzel");
            Factory2D catFactory2D1 = hsp_catiaProfil.OpenEdition();

            //Punkte 
            Point2D catP2D_Ursprung = catFactory2D1.CreatePoint(x0, y0);

            Point2D catP2D_StartPkt_Fußkreis = catFactory2D1.CreatePoint(StartPkt_Fußkreis_x, StartPkt_Fußkreis_y);
            Point2D catP2D_SP_FußkreisVerrundungsRadius1 = catFactory2D1.CreatePoint(SP_FußkreisVerrundungsRadius_x, SP_FußkreisVerrundungsRadius_y);
            Point2D catP2D_SP_FußkreisVerrundungsRadius2 = catFactory2D1.CreatePoint(-SP_FußkreisVerrundungsRadius_x, SP_FußkreisVerrundungsRadius_y);

            Point2D catP2D_MP_EvolventenKreis1 = catFactory2D1.CreatePoint(MP_EvolventenKreis_x, MP_EvolventenKreis_y);
            Point2D catP2D_MP_EvolventenKreis2 = catFactory2D1.CreatePoint(-MP_EvolventenKreis_x, MP_EvolventenKreis_y);

            Point2D catP2D_SP_EvolventenKopfKreis1 = catFactory2D1.CreatePoint(SP_EvolventenKopfKreis_x, SP_EvolventenKopfKreis_y);
            Point2D catP2D_SP_EvolventenKopfKreis2 = catFactory2D1.CreatePoint(-SP_EvolventenKopfKreis_x, SP_EvolventenKopfKreis_y);

            Point2D catP2D_MP_Verrundung1 = catFactory2D1.CreatePoint(MP_Verrundung_x, MP_Verrundung_y);
            Point2D catP2D_MP_Verrundung2 = catFactory2D1.CreatePoint(-MP_Verrundung_x, MP_Verrundung_y);

            Point2D catP2D_SP_EvolventeVerrundung1 = catFactory2D1.CreatePoint(SP_EvolventeVerrundung_x, SP_EvolventeVerrundung_y);
            Point2D catP2D_SP_EvolventeVerrundung2 = catFactory2D1.CreatePoint(-SP_EvolventeVerrundung_x, SP_EvolventeVerrundung_y); 


            //Kreise
            Circle2D catC2D_Frußkreis = catFactory2D1.CreateCircle(x0, y0, df_r, 0, 0);
            catC2D_Frußkreis.CenterPoint = catP2D_Ursprung;
            catC2D_Frußkreis.StartPoint = catP2D_SP_FußkreisVerrundungsRadius1;
            catC2D_Frußkreis.EndPoint = catP2D_StartPkt_Fußkreis;

            Circle2D catC2D_Kopfkreis = catFactory2D1.CreateCircle(x0, y0, da_r, 0, 0);
            catC2D_Kopfkreis.CenterPoint = catP2D_Ursprung;
            catC2D_Kopfkreis.StartPoint = catP2D_SP_EvolventenKopfKreis2;
            catC2D_Kopfkreis.EndPoint = catP2D_SP_EvolventenKopfKreis1;

            Circle2D catC2D_EvolventenKreis1 = catFactory2D1.CreateCircle(MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r, 0, 0);
            catC2D_EvolventenKreis1.CenterPoint = catP2D_MP_EvolventenKreis1;
            catC2D_EvolventenKreis1.StartPoint = catP2D_SP_EvolventenKopfKreis1;
            catC2D_EvolventenKreis1.EndPoint = catP2D_SP_EvolventeVerrundung1;

            Circle2D catC2D_Evolventenkreis2 = catFactory2D1.CreateCircle(-MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r, 0, 0);
            catC2D_Evolventenkreis2.CenterPoint = catP2D_MP_EvolventenKreis2;
            catC2D_Evolventenkreis2.StartPoint = catP2D_SP_EvolventeVerrundung2;
            catC2D_Evolventenkreis2.EndPoint = catP2D_SP_EvolventenKopfKreis2;

            Circle2D catC2D_VerrundungsKreis1 = catFactory2D1.CreateCircle(MP_Verrundung_x, MP_Verrundung_y, vd_r, 0, 0);
            catC2D_VerrundungsKreis1.CenterPoint = catP2D_MP_Verrundung1;
            catC2D_VerrundungsKreis1.StartPoint = catP2D_SP_FußkreisVerrundungsRadius1;
            catC2D_VerrundungsKreis1.EndPoint = catP2D_SP_EvolventeVerrundung1;

            Circle2D catC2D_VerrundungsKreis2 = catFactory2D1.CreateCircle(-MP_Verrundung_x, MP_Verrundung_y, vd_r, 0, 0);
            catC2D_VerrundungsKreis2.CenterPoint = catP2D_MP_Verrundung2;
            catC2D_VerrundungsKreis2.StartPoint = catP2D_SP_EvolventeVerrundung2;
            catC2D_VerrundungsKreis2.EndPoint = catP2D_SP_FußkreisVerrundungsRadius2;

            hsp_catiaProfil.CloseEdition();

            hsp_catiaPart.Part.Update();
        }

        public void ErzeugeKreismusterAußenverzahnung(Data dat)
        {
            ShapeFactory SF = (ShapeFactory)hsp_catiaPart.Part.ShapeFactory;
            HybridShapeFactory HSF = (HybridShapeFactory)hsp_catiaPart.Part.HybridShapeFactory;

            //Skizze und Referenzen
            Factory2D Factory2D1 = hsp_catiaProfil.Factory2D;

            HybridShapePointCoord Ursprung = HSF.AddNewPointCoord(0, 0, 0);
            Reference RefUrsprung = hsp_catiaPart.Part.CreateReferenceFromObject(Ursprung);
            HybridShapeDirection XDir = HSF.AddNewDirectionByCoord(1, 0, 0);
            Reference RefXDir = hsp_catiaPart.Part.CreateReferenceFromObject(XDir);

            //Kreismuster mit Daten füllen
            CircPattern Kreismuster = SF.AddNewSurfacicCircPattern(Factory2D1, 1, 2, 0, 0, 1, 1, RefUrsprung, RefXDir, false, 0, true, false);
            Kreismuster.CircularPatternParameters = CatCircularPatternParameters.catInstancesandAngularSpacing;
            AngularRepartition angularRepartition1 = Kreismuster.AngularRepartition;
            Angle angle1 = angularRepartition1.AngularSpacing;
            angle1.Value = Convert.ToDouble(360 / dat.getZaehnezahlZahnrad1());
            AngularRepartition angularRepartition2 = Kreismuster.AngularRepartition;
            IntParam intParam1 = angularRepartition2.InstancesCount;
            intParam1.Value = Convert.ToInt32(dat.getZaehnezahlZahnrad1()) + 1;

            //geschlossene Kontur
            Reference Ref_Kreismuster = hsp_catiaPart.Part.CreateReferenceFromObject(Kreismuster);
            HybridShapeAssemble Verbindung = HSF.AddNewJoin(Ref_Kreismuster, Ref_Kreismuster);
            Reference Ref_Verbindung = hsp_catiaPart.Part.CreateReferenceFromObject(Verbindung);

            HSF.GSMVisibility(Ref_Verbindung, 0);

            hsp_catiaPart.Part.Update();

            Bodies bodies = hsp_catiaPart.Part.Bodies;
            Body myBody = bodies.Add();
            myBody.set_Name("Zahnrad");
            myBody.InsertHybridShape(Verbindung);

            hsp_catiaPart.Part.Update();

            //Erzeuge Block aus Skizze
            hsp_catiaPart.Part.InWorkObject = myBody;
            Pad myPad = SF.AddNewPadFromRef(Ref_Verbindung, dat.getBreiteZahnrad1());
            
            hsp_catiaPart.Part.Update();
        }

        public void ErzeugeKreismusterMitBohrung(Data dat)
        {
             //Erzeuge Kreismuster

             ShapeFactory SF = (ShapeFactory)hsp_catiaPart.Part.ShapeFactory;
             HybridShapeFactory HSF = (HybridShapeFactory)hsp_catiaPart.Part.HybridShapeFactory;

             //Skizze und Referenzen
             Factory2D Factory2D1 = hsp_catiaProfil.Factory2D;

             HybridShapePointCoord Ursprung = HSF.AddNewPointCoord(0, 0, 0);
             Reference RefUrsprung = hsp_catiaPart.Part.CreateReferenceFromObject(Ursprung);
             HybridShapeDirection XDir = HSF.AddNewDirectionByCoord(1, 0, 0);
             Reference RefXDir = hsp_catiaPart.Part.CreateReferenceFromObject(XDir);

             //Kreismuster mit Daten füllen
             CircPattern Kreismuster = SF.AddNewSurfacicCircPattern(Factory2D1, 1, 2, 0, 0, 1, 1, RefUrsprung, RefXDir, false, 0, true, false);
             Kreismuster.CircularPatternParameters = CatCircularPatternParameters.catInstancesandAngularSpacing;
             AngularRepartition angularRepartition1 = Kreismuster.AngularRepartition;
             Angle angle1 = angularRepartition1.AngularSpacing;
             angle1.Value = Convert.ToDouble(360 / dat.getZaehnezahlZahnrad1());
             AngularRepartition angularRepartition2 = Kreismuster.AngularRepartition;
             IntParam intParam1 = angularRepartition2.InstancesCount;
             intParam1.Value = Convert.ToInt32(dat.getZaehnezahlZahnrad1()) + 1;

             //geschlossene Kontur
             Reference Ref_Kreismuster = hsp_catiaPart.Part.CreateReferenceFromObject(Kreismuster);
             HybridShapeAssemble Verbindung = HSF.AddNewJoin(Ref_Kreismuster, Ref_Kreismuster);
             Reference Ref_Verbindung = hsp_catiaPart.Part.CreateReferenceFromObject(Verbindung);

             HSF.GSMVisibility(Ref_Verbindung, 0);

             hsp_catiaPart.Part.Update();

             Bodies bodies = hsp_catiaPart.Part.Bodies;
             Body myBody = bodies.Add();
             myBody.set_Name("Zahnrad");
             myBody.InsertHybridShape(Verbindung);

             hsp_catiaPart.Part.Update();

             //Erzeuge Block aus Skizze
             hsp_catiaPart.Part.InWorkObject = myBody;
             Pad myPad = SF.AddNewPadFromRef(Ref_Verbindung, dat.getBreiteZahnrad1());

             hsp_catiaPart.Part.Update();


            //Erzeuge Skizze für Bohrung          
            Reference RefBohrung1 = hsp_catiaPart.Part.CreateReferenceFromBRepName("FSur:(Face:(Brp:(Pad.1;2);None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", myPad);
            Hole catBohrung1 = SF.AddNewHoleFromPoint(0, 0, 0, RefBohrung1, dat.getBreiteZahnrad1());
            Length catLengthBohrung1 = catBohrung1.Diameter;
            catLengthBohrung1.Value = Convert.ToDouble(dat.getTeilkreisdurchmesserZahnrad1() / 2);

            hsp_catiaPart.Part.Update();
        }



        //InnenVerzahnung
        public void ErstelleProfilInnen(Data dat)
        {
            //geometrisches set auswählen und umbenennen
            HybridBodies catHybridBodies_I = hsp_catiaPart.Part.HybridBodies;
            HybridBody catHybridBody_I;

            try
            {
                catHybridBody_I = catHybridBodies_I.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                MessageBox.Show("Kein geometrisches Set gefunden!\nEin PART manuell erzeugen und darauf achten, dass ein 'Geometisches Set' aktiviert ist.", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            catHybridBody_I.set_Name("Profile");

            //Neue Skizze im ausgewählten geometrischen Set anlegen
            Sketches catSketches_I = catHybridBody_I.HybridSketches;
            OriginElements catOriginElements_I = hsp_catiaPart.Part.OriginElements;
            Reference catReference_I = (Reference)catOriginElements_I.PlaneYZ;
            hsp_catiaProfil = catSketches_I.Add(catReference_I);

            //Achsensystem in Skizze erzeugen
            ErzeugeAchsensystem();

            //Part aktualisieren
            hsp_catiaPart.Part.Update();

            hsp_catiaProfil.set_Name("InnenverzahnungBlock");
            Factory2D catFactory_I = hsp_catiaProfil.OpenEdition();

            Circle2D catC2D_I = catFactory_I.CreateClosedCircle(0, 0, dat.getFußkreisdurchmesser_iZahnrad1());

            hsp_catiaProfil.CloseEdition();
            hsp_catiaPart.Part.Update();

            ShapeFactory SF_I = (ShapeFactory)hsp_catiaPart.Part.ShapeFactory;
            HybridShapeFactory HSF_I = (HybridShapeFactory)hsp_catiaPart.Part.HybridShapeFactory;

            //Erzeuge Block aus Skizze
            hsp_catiaPart.Part.InWorkObject = hsp_catiaPart.Part.MainBody;
            Pad myPad = SF_I.AddNewPad(hsp_catiaProfil, dat.getBreiteZahnrad1());

            hsp_catiaPart.Part.Update();



            //Neue Skizze im ausgewählten geometrischen Set anlegen
            Sketches catSketches1 = catHybridBody_I.HybridSketches;
            OriginElements catOriginElements = hsp_catiaPart.Part.OriginElements;
            Reference catReference1 = (Reference)catOriginElements.PlaneYZ;
            hsp_catiaProfil = catSketches1.Add(catReference1);

            //Achsensystem in Skizze erzeugen
            ErzeugeAchsensystem();

            //Part aktualisieren
            hsp_catiaPart.Part.Update();


            //HilfsRadien
            double d_r = (dat.getModulZahnrad1() * dat.getZaehnezahlZahnrad1()) / 2;
            double hk_r = d_r * 1.06;
            double da_r = d_r - (1.25 * dat.getModulZahnrad1());
            double df_r = d_r + dat.getModulZahnrad1();
            double vd_r = 0.35 * dat.getModulZahnrad1();

            //HilfsWinkel
            double alpha = 20;
            double beta = 90 / dat.getZaehnezahlZahnrad1();
            double beta_r = Math.PI * beta / 180;
            double gamma = 90 - (alpha - beta);
            double gamma_r = Math.PI * gamma / 180;
            double ta = 360.0 / dat.getZaehnezahlZahnrad1();
            double ta_r = Math.PI * ta / 180;

            //Nullpunkte
            double x0 = 0;
            double y0 = 0;

            //MittelPunkt EvolventenKreis
            double MP_EvolventenKreis_x = hk_r * Math.Cos(gamma_r);
            double MP_EvolventenKreis_y = hk_r * Math.Sin(gamma_r);

            // SchnittPunkt Evolventenkreis & Teilkreisradius
            double SP_EvolventenTeilKreis_x = -d_r * Math.Sin(beta_r);
            double SP_EvolventenTeilKreis_y = d_r * Math.Cos(beta_r);

            //Evolventenkreis Radius
            double Evolventenkreis_r = Math.Sqrt(Math.Pow((MP_EvolventenKreis_x - SP_EvolventenTeilKreis_x), 2) + Math.Pow((MP_EvolventenKreis_y - SP_EvolventenTeilKreis_y), 2));

            //SchnittPunkt Evolventenkreis & Kopfkreisradius
            double SP_EvolventenKopfKreis_x = Schnittpunkt_x(x0, y0, da_r, MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r);
            double SP_EvolventenKopfKreis_y = Schnittpunkt_y(x0, y0, da_r, MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r);

            //MittelPunkt VerrundungsRadius
            double MP_Verrundung_x = Schnittpunkt_x(x0, y0, df_r + vd_r, MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r + vd_r);
            double MP_Verrundung_y = Schnittpunkt_y(x0, y0, df_r + vd_r, MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r + vd_r);

            //SchnittPunkt Evolventenkreis & Verrundungsradius
            double SP_EvolventeVerrundung_x = Schnittpunkt_x(MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r, MP_Verrundung_x, MP_Verrundung_y, vd_r);
            double SP_EvolventeVerrundung_y = Schnittpunkt_y(MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r, MP_Verrundung_x, MP_Verrundung_y, vd_r);

            //SchnittPunkt Fußkreis & Verrundungs Radius
            double SP_FußkreisVerrundungsRadius_x = Schnittpunkt_x(x0, y0, df_r, MP_Verrundung_x, MP_Verrundung_y, vd_r);
            double SP_FußkreisVerrundungsRadius_y = Schnittpunkt_y(x0, y0, df_r, MP_Verrundung_x, MP_Verrundung_y, vd_r);

            //StartPunkt Fußkreis Radius
            double phi = ta_r - Math.Atan(Math.Abs(SP_FußkreisVerrundungsRadius_x) / Math.Abs(SP_FußkreisVerrundungsRadius_y));
            double StartPkt_Fußkreis_x = -df_r * Math.Sin(phi);
            double StartPkt_Fußkreis_y = df_r * Math.Cos(phi);

            //Skizze umbenennen und öffnen
            hsp_catiaProfil.set_Name("InnenverzahnungEinzel");
            Factory2D catFactory2D1 = hsp_catiaProfil.OpenEdition();

            //Punkte 
            Point2D catP2D_Ursprung = catFactory2D1.CreatePoint(x0, y0);

            Point2D catP2D_StartPkt_Fußkreis = catFactory2D1.CreatePoint(StartPkt_Fußkreis_x, StartPkt_Fußkreis_y);
            Point2D catP2D_SP_FußkreisVerrundungsRadius1 = catFactory2D1.CreatePoint(SP_FußkreisVerrundungsRadius_x, SP_FußkreisVerrundungsRadius_y);
            Point2D catP2D_SP_FußkreisVerrundungsRadius2 = catFactory2D1.CreatePoint(-SP_FußkreisVerrundungsRadius_x, SP_FußkreisVerrundungsRadius_y);

            Point2D catP2D_MP_EvolventenKreis1 = catFactory2D1.CreatePoint(MP_EvolventenKreis_x, MP_EvolventenKreis_y);
            Point2D catP2D_MP_EvolventenKreis2 = catFactory2D1.CreatePoint(-MP_EvolventenKreis_x, MP_EvolventenKreis_y);

            Point2D catP2D_SP_EvolventenKopfKreis1 = catFactory2D1.CreatePoint(SP_EvolventenKopfKreis_x, SP_EvolventenKopfKreis_y);
            Point2D catP2D_SP_EvolventenKopfKreis2 = catFactory2D1.CreatePoint(-SP_EvolventenKopfKreis_x, SP_EvolventenKopfKreis_y);

            Point2D catP2D_MP_Verrundung1 = catFactory2D1.CreatePoint(MP_Verrundung_x, MP_Verrundung_y);
            Point2D catP2D_MP_Verrundung2 = catFactory2D1.CreatePoint(-MP_Verrundung_x, MP_Verrundung_y);

            Point2D catP2D_SP_EvolventeVerrundung1 = catFactory2D1.CreatePoint(SP_EvolventeVerrundung_x, SP_EvolventeVerrundung_y);
            Point2D catP2D_SP_EvolventeVerrundung2 = catFactory2D1.CreatePoint(-SP_EvolventeVerrundung_x, SP_EvolventeVerrundung_y);


            //Kreise
            Circle2D catC2D_Frußkreis = catFactory2D1.CreateCircle(x0, y0, df_r, 0, 0);
            catC2D_Frußkreis.CenterPoint = catP2D_Ursprung;
            catC2D_Frußkreis.StartPoint = catP2D_SP_FußkreisVerrundungsRadius1;
            catC2D_Frußkreis.EndPoint = catP2D_StartPkt_Fußkreis;

            Circle2D catC2D_Kopfkreis = catFactory2D1.CreateCircle(x0, y0, da_r, 0, 0);
            catC2D_Kopfkreis.CenterPoint = catP2D_Ursprung;
            catC2D_Kopfkreis.StartPoint = catP2D_SP_EvolventenKopfKreis2;
            catC2D_Kopfkreis.EndPoint = catP2D_SP_EvolventenKopfKreis1;

            Circle2D catC2D_EvolventenKreis1 = catFactory2D1.CreateCircle(MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r, 0, 0);
            catC2D_EvolventenKreis1.CenterPoint = catP2D_MP_EvolventenKreis1;
            catC2D_EvolventenKreis1.StartPoint = catP2D_SP_EvolventenKopfKreis1;
            catC2D_EvolventenKreis1.EndPoint = catP2D_SP_EvolventeVerrundung1;

            Circle2D catC2D_Evolventenkreis2 = catFactory2D1.CreateCircle(-MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r, 0, 0);
            catC2D_Evolventenkreis2.CenterPoint = catP2D_MP_EvolventenKreis2;
            catC2D_Evolventenkreis2.StartPoint = catP2D_SP_EvolventeVerrundung2;
            catC2D_Evolventenkreis2.EndPoint = catP2D_SP_EvolventenKopfKreis2;

            Circle2D catC2D_VerrundungsKreis1 = catFactory2D1.CreateCircle(MP_Verrundung_x, MP_Verrundung_y, vd_r, 0, 0);
            catC2D_VerrundungsKreis1.CenterPoint = catP2D_MP_Verrundung1;
            catC2D_VerrundungsKreis1.StartPoint = catP2D_SP_FußkreisVerrundungsRadius1;
            catC2D_VerrundungsKreis1.EndPoint = catP2D_SP_EvolventeVerrundung1;

            Circle2D catC2D_VerrundungsKreis2 = catFactory2D1.CreateCircle(-MP_Verrundung_x, MP_Verrundung_y, vd_r, 0, 0);
            catC2D_VerrundungsKreis2.CenterPoint = catP2D_MP_Verrundung2;
            catC2D_VerrundungsKreis2.StartPoint = catP2D_SP_EvolventeVerrundung2;
            catC2D_VerrundungsKreis2.EndPoint = catP2D_SP_FußkreisVerrundungsRadius2;

            hsp_catiaProfil.CloseEdition();

            hsp_catiaPart.Part.Update();

            //Skizze und Referenzen
            Factory2D Factory2D1 = hsp_catiaProfil.Factory2D;

            HybridShapePointCoord Ursprung = HSF_I.AddNewPointCoord(0, 0, 0);
            Reference RefUrsprung = hsp_catiaPart.Part.CreateReferenceFromObject(Ursprung);
            HybridShapeDirection XDir = HSF_I.AddNewDirectionByCoord(1, 0, 0);
            Reference RefXDir = hsp_catiaPart.Part.CreateReferenceFromObject(XDir);

            //Kreismuster mit Daten füllen
            CircPattern Kreismuster = SF_I.AddNewSurfacicCircPattern(Factory2D1, 1, 2, 0, 0, 1, 1, RefUrsprung, RefXDir, false, 0, true, false);
            Kreismuster.CircularPatternParameters = CatCircularPatternParameters.catInstancesandAngularSpacing;
            AngularRepartition angularRepartition1 = Kreismuster.AngularRepartition;
            Angle angle1 = angularRepartition1.AngularSpacing;
            angle1.Value = Convert.ToDouble(360 / dat.getZaehnezahlZahnrad1());
            AngularRepartition angularRepartition2 = Kreismuster.AngularRepartition;
            IntParam intParam1 = angularRepartition2.InstancesCount;
            intParam1.Value = Convert.ToInt32(dat.getZaehnezahlZahnrad1()) + 1;

            //geschlossene Kontur
            Reference Ref_Kreismuster = hsp_catiaPart.Part.CreateReferenceFromObject(Kreismuster);
            HybridShapeAssemble Verbindung = HSF_I.AddNewJoin(Ref_Kreismuster, Ref_Kreismuster);
            Reference Ref_Verbindung = hsp_catiaPart.Part.CreateReferenceFromObject(Verbindung);

            HSF_I.GSMVisibility(Ref_Verbindung, 0);

            hsp_catiaPart.Part.Update();

            /*Bodies bodies = hsp_catiaPart.Part.Bodies;
            Body myBody = bodies.Add();
            myBody.set_Name("Zahnrad");
            myBody.InsertHybridShape(Verbindung);

            hsp_catiaPart.Part.Update();*/

            //Tasche für Innenverzahnung(grob)
            hsp_catiaPart.Part.InWorkObject = hsp_catiaPart.Part.MainBody;

            Pocket catPocketInnen = SF_I.AddNewPocketFromRef(Ref_Verbindung, dat.getBreiteZahnrad1());
            hsp_catiaPart.Part.Update();

        }

        /*public void ErzeugeKreismusterInnen(Data dat)
        {
            ShapeFactory SF = (ShapeFactory)hsp_catiaPart.Part.ShapeFactory;
            HybridShapeFactory HSF = (HybridShapeFactory)hsp_catiaPart.Part.HybridShapeFactory;

            //Skizze und Referenzen
            Factory2D Factory2D1 = hsp_catiaProfil.Factory2D;

            HybridShapePointCoord Ursprung = HSF.AddNewPointCoord(0, 0, 0);
            Reference RefUrsprung = hsp_catiaPart.Part.CreateReferenceFromObject(Ursprung);
            HybridShapeDirection XDir = HSF.AddNewDirectionByCoord(1, 0, 0);
            Reference RefXDir = hsp_catiaPart.Part.CreateReferenceFromObject(XDir);

            //Kreismuster mit Daten füllen
            CircPattern Kreismuster = SF.AddNewSurfacicCircPattern(Factory2D1, 1, 2, 0, 0, 1, 1, RefUrsprung, RefXDir, false, 0, true, false);
            Kreismuster.CircularPatternParameters = CatCircularPatternParameters.catInstancesandAngularSpacing;
            AngularRepartition angularRepartition1 = Kreismuster.AngularRepartition;
            Angle angle1 = angularRepartition1.AngularSpacing;
            angle1.Value = Convert.ToDouble(360 / dat.getZaehnezahlZahnrad1());
            AngularRepartition angularRepartition2 = Kreismuster.AngularRepartition;
            IntParam intParam1 = angularRepartition2.InstancesCount;
            intParam1.Value = Convert.ToInt32(dat.getZaehnezahlZahnrad1()) + 1;

            //geschlossene Kontur
            Reference Ref_Kreismuster = hsp_catiaPart.Part.CreateReferenceFromObject(Kreismuster);
            HybridShapeAssemble Verbindung = HSF.AddNewJoin(Ref_Kreismuster, Ref_Kreismuster);
            Reference Ref_Verbindung = hsp_catiaPart.Part.CreateReferenceFromObject(Verbindung);

            HSF.GSMVisibility(Ref_Verbindung, 0);

            hsp_catiaPart.Part.Update();

            Bodies bodies = hsp_catiaPart.Part.Bodies;
            Body myBody = bodies.Add();
            myBody.set_Name("Zahnrad");
            myBody.InsertHybridShape(Verbindung);

            hsp_catiaPart.Part.Update();

            //Tasche für Innenverzahnung(grob)
            hsp_catiaPart.Part.InWorkObject = hsp_catiaPart.Part.MainBody;
           
            Pocket catPocketInnen = SF.AddNewPocketFromRef(Ref_Verbindung, dat.getBreiteZahnrad1());
            hsp_catiaPart.Part.Update();
        }*/
        /*public void AußenverzahnungGegenrad(Data dat)
        {
            //HilfsRadien
            double d_r = (dat.getModulZahnrad2() * dat.getZaehnezahlZahnrad2()) / 2;
            double hk_r = d_r * 0.94;
            double df_r = d_r - (1.25 * dat.getModulZahnrad2());
            double da_r = d_r + dat.getModulZahnrad2();
            double vd_r = 0.35 * dat.getModulZahnrad2();

            //HilfsWinkel
            double alpha = 20;
            double beta = 90 / dat.getZaehnezahlZahnrad2();
            double beta_r = Math.PI * beta / 180;
            double gamma = 90 - (alpha - beta);
            double gamma_r = Math.PI * gamma / 180;
            double ta = 360.0 / dat.getZaehnezahlZahnrad2();
            double ta_r = Math.PI * ta / 180;

            //Nullpunkte
            double x0 = dat.getAchsabstand_a();
            double y0 = 0;

            //MittelPunkt EvolventenKreis
            double MP_EvolventenKreis_x = hk_r * Math.Cos(gamma_r);
            double MP_EvolventenKreis_y = hk_r * Math.Sin(gamma_r);

            // SchnittPunkt Evolventenkreis & Teilkreisradius
            double SP_EvolventenTeilKreis_x = -d_r * Math.Sin(beta_r);
            double SP_EvolventenTeilKreis_y = d_r * Math.Cos(beta_r);

            //Evolventenkreis Radius
            double Evolventenkreis_r = Math.Sqrt(Math.Pow((MP_EvolventenKreis_x - SP_EvolventenTeilKreis_x), 2) + Math.Pow((MP_EvolventenKreis_y - SP_EvolventenTeilKreis_y), 2));

            //SchnittPunkt Evolventenkreis & Kopfkreisradius
            double SP_EvolventenKopfKreis_x = Schnittpunkt_x(x0, y0, da_r, MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r);
            double SP_EvolventenKopfKreis_y = Schnittpunkt_y(x0, y0, da_r, MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r);

            //MittelPunkt VerrundungsRadius
            double MP_Verrundung_x = Schnittpunkt_x(x0, y0, df_r + vd_r, MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r + vd_r);
            double MP_Verrundung_y = Schnittpunkt_y(x0, y0, df_r + vd_r, MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r + vd_r);

            //SchnittPunkt Evolventenkreis & Verrundungsradius
            double SP_EvolventeVerrundung_x = Schnittpunkt_x(MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r, MP_Verrundung_x, MP_Verrundung_y, vd_r);
            double SP_EvolventeVerrundung_y = Schnittpunkt_y(MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r, MP_Verrundung_x, MP_Verrundung_y, vd_r);

            //SchnittPunkt Fußkreis & Verrundungs Radius
            double SP_FußkreisVerrundungsRadius_x = Schnittpunkt_x(x0, y0, df_r, MP_Verrundung_x, MP_Verrundung_y, vd_r);
            double SP_FußkreisVerrundungsRadius_y = Schnittpunkt_y(x0, y0, df_r, MP_Verrundung_x, MP_Verrundung_y, vd_r);

            //StartPunkt Fußkreis Radius
            double phi = ta_r - Math.Atan(Math.Abs(SP_FußkreisVerrundungsRadius_x) / Math.Abs(SP_FußkreisVerrundungsRadius_y));
            double StartPkt_Fußkreis_x = -df_r * Math.Sin(phi);
            double StartPkt_Fußkreis_y = df_r * Math.Cos(phi);

            //Skizze umbenennen und öffnen
            hsp_catiaProfil.set_Name("AußenverzahnungEinzel");
            Factory2D catFactory2D1 = hsp_catiaProfil.OpenEdition();

            //Punkte 
            Point2D catP2D_Ursprung = catFactory2D1.CreatePoint(x0, y0);

            Point2D catP2D_StartPkt_Fußkreis = catFactory2D1.CreatePoint(StartPkt_Fußkreis_x, StartPkt_Fußkreis_y);
            Point2D catP2D_SP_FußkreisVerrundungsRadius1 = catFactory2D1.CreatePoint(SP_FußkreisVerrundungsRadius_x, SP_FußkreisVerrundungsRadius_y);
            Point2D catP2D_SP_FußkreisVerrundungsRadius2 = catFactory2D1.CreatePoint(-SP_FußkreisVerrundungsRadius_x, SP_FußkreisVerrundungsRadius_y);

            Point2D catP2D_MP_EvolventenKreis1 = catFactory2D1.CreatePoint(MP_EvolventenKreis_x, MP_EvolventenKreis_y);
            Point2D catP2D_MP_EvolventenKreis2 = catFactory2D1.CreatePoint(-MP_EvolventenKreis_x, MP_EvolventenKreis_y);

            Point2D catP2D_SP_EvolventenKopfKreis1 = catFactory2D1.CreatePoint(SP_EvolventenKopfKreis_x, SP_EvolventenKopfKreis_y);
            Point2D catP2D_SP_EvolventenKopfKreis2 = catFactory2D1.CreatePoint(-SP_EvolventenKopfKreis_x, SP_EvolventenKopfKreis_y);

            Point2D catP2D_MP_Verrundung1 = catFactory2D1.CreatePoint(MP_Verrundung_x, MP_Verrundung_y);
            Point2D catP2D_MP_Verrundung2 = catFactory2D1.CreatePoint(-MP_Verrundung_x, MP_Verrundung_y);

            Point2D catP2D_SP_EvolventeVerrundung1 = catFactory2D1.CreatePoint(SP_EvolventeVerrundung_x, SP_EvolventeVerrundung_y);
            Point2D catP2D_SP_EvolventeVerrundung2 = catFactory2D1.CreatePoint(-SP_EvolventeVerrundung_x, SP_EvolventeVerrundung_y);


            //Kreise
            Circle2D catC2D_Frußkreis = catFactory2D1.CreateCircle(x0, y0, df_r, 0, 0);
            catC2D_Frußkreis.CenterPoint = catP2D_Ursprung;
            catC2D_Frußkreis.StartPoint = catP2D_SP_FußkreisVerrundungsRadius1;
            catC2D_Frußkreis.EndPoint = catP2D_StartPkt_Fußkreis;

            Circle2D catC2D_Kopfkreis = catFactory2D1.CreateCircle(x0, y0, da_r, 0, 0);
            catC2D_Kopfkreis.CenterPoint = catP2D_Ursprung;
            catC2D_Kopfkreis.StartPoint = catP2D_SP_EvolventenKopfKreis2;
            catC2D_Kopfkreis.EndPoint = catP2D_SP_EvolventenKopfKreis1;

            Circle2D catC2D_EvolventenKreis1 = catFactory2D1.CreateCircle(MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r, 0, 0);
            catC2D_EvolventenKreis1.CenterPoint = catP2D_MP_EvolventenKreis1;
            catC2D_EvolventenKreis1.StartPoint = catP2D_SP_EvolventenKopfKreis1;
            catC2D_EvolventenKreis1.EndPoint = catP2D_SP_EvolventeVerrundung1;

            Circle2D catC2D_Evolventenkreis2 = catFactory2D1.CreateCircle(-MP_EvolventenKreis_x, MP_EvolventenKreis_y, Evolventenkreis_r, 0, 0);
            catC2D_Evolventenkreis2.CenterPoint = catP2D_MP_EvolventenKreis2;
            catC2D_Evolventenkreis2.StartPoint = catP2D_SP_EvolventeVerrundung2;
            catC2D_Evolventenkreis2.EndPoint = catP2D_SP_EvolventenKopfKreis2;

            Circle2D catC2D_VerrundungsKreis1 = catFactory2D1.CreateCircle(MP_Verrundung_x, MP_Verrundung_y, vd_r, 0, 0);
            catC2D_VerrundungsKreis1.CenterPoint = catP2D_MP_Verrundung1;
            catC2D_VerrundungsKreis1.StartPoint = catP2D_SP_FußkreisVerrundungsRadius1;
            catC2D_VerrundungsKreis1.EndPoint = catP2D_SP_EvolventeVerrundung1;

            Circle2D catC2D_VerrundungsKreis2 = catFactory2D1.CreateCircle(-MP_Verrundung_x, MP_Verrundung_y, vd_r, 0, 0);
            catC2D_VerrundungsKreis2.CenterPoint = catP2D_MP_Verrundung2;
            catC2D_VerrundungsKreis2.StartPoint = catP2D_SP_EvolventeVerrundung2;
            catC2D_VerrundungsKreis2.EndPoint = catP2D_SP_FußkreisVerrundungsRadius2;

            hsp_catiaProfil.CloseEdition();

            hsp_catiaPart.Part.Update();
        }*/

        /*public void ErzeugeKreismusterGegenrad(Data dat)
        {
            ShapeFactory SF = (ShapeFactory)hsp_catiaPart.Part.ShapeFactory;
            HybridShapeFactory HSF = (HybridShapeFactory)hsp_catiaPart.Part.HybridShapeFactory;

            //Skizze und Referenzen
            Factory2D Factory2D1 = hsp_catiaProfil.Factory2D;

            HybridShapePointCoord Ursprung = HSF.AddNewPointCoord(dat.getAchsabstand_a(), 0, 0);
            Reference RefUrsprung = hsp_catiaPart.Part.CreateReferenceFromObject(Ursprung);
            HybridShapeDirection XDir = HSF.AddNewDirectionByCoord(1, 0, 0);
            Reference RefXDir = hsp_catiaPart.Part.CreateReferenceFromObject(XDir);

            //Kreismuster mit Daten füllen
            CircPattern Kreismuster = SF.AddNewSurfacicCircPattern(Factory2D1, 1, 2, 0, 0, 1, 1, RefUrsprung, RefXDir, false, 0, true, false);
            Kreismuster.CircularPatternParameters = CatCircularPatternParameters.catInstancesandAngularSpacing;
            AngularRepartition angularRepartition1 = Kreismuster.AngularRepartition;
            Angle angle1 = angularRepartition1.AngularSpacing;
            angle1.Value = Convert.ToDouble(360 / dat.getZaehnezahlZahnrad2());
            AngularRepartition angularRepartition2 = Kreismuster.AngularRepartition;
            IntParam intParam1 = angularRepartition2.InstancesCount;
            intParam1.Value = Convert.ToInt32(dat.getZaehnezahlZahnrad2()) + 1;

            //geschlossene Kontur
            Reference Ref_Kreismuster = hsp_catiaPart.Part.CreateReferenceFromObject(Kreismuster);
            HybridShapeAssemble Verbindung = HSF.AddNewJoin(Ref_Kreismuster, Ref_Kreismuster);
            Reference Ref_Verbindung = hsp_catiaPart.Part.CreateReferenceFromObject(Verbindung);

            HSF.GSMVisibility(Ref_Verbindung, 0);

            hsp_catiaPart.Part.Update();

            Bodies bodies = hsp_catiaPart.Part.Bodies;
            Body myBody = bodies.Add();
            myBody.set_Name("Gegenrad");
            myBody.InsertHybridShape(Verbindung);

            hsp_catiaPart.Part.Update();

            //Erzeuge Block aus Skizze
            hsp_catiaPart.Part.InWorkObject = myBody;
            Pad myPad = SF.AddNewPadFromRef(Ref_Verbindung, dat.getBreiteZahnrad2());

            hsp_catiaPart.Part.Update();
        }*/

        //Berechnungen der SchnittPunkte
        private double Schnittpunkt_x(double MP1_x, double MP1_y, double r1, double MP2_x, double MP2_y, double r2)
        {
            double d = Math.Sqrt(Math.Pow((MP1_x - MP2_x), 2) + Math.Pow((MP1_y - MP2_y), 2));
            double l = (Math.Pow(r1, 2) - Math.Pow(r2, 2) + Math.Pow(d, 2)) / (d * 2);
            double h;
            double ii = 0.00001;

            if (r1 - l < -ii)
            {
                MessageBox.Show("Bitte überprüfen Sie die Eingangsparameter.", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            if (Math.Abs(r1 - l) < ii)
            {
                h = 0;
            }
            else
            {
                h = Math.Sqrt(Math.Pow(r1, 2) - Math.Pow(l, 2));
            }

            return l * (MP2_x - MP1_x) / d - h * (MP2_y - MP1_y) / d + MP1_x;
        }
        private double Schnittpunkt_y(double MP1_x, double MP1_y, double r1, double MP2_x, double MP2_y, double r2)
        {
            double d = Math.Sqrt(Math.Pow((MP1_x - MP2_x), 2) + Math.Pow((MP1_y - MP2_y), 2));
            double l = (Math.Pow(r1, 2) - Math.Pow(r2, 2) + Math.Pow(d, 2)) / (d * 2);
            double h;
            double ii = 0.00001;

            if (r1 - l < -ii)
            {
                MessageBox.Show("Bitte überprüfen Sie die Eingangsparameter.", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            if (Math.Abs(r1 - l) < ii)
            {
                h = 0;
            }
            else
            {
                h = Math.Sqrt(Math.Pow(r1, 2) - Math.Pow(l, 2));
            }

            return l * (MP2_y - MP1_y) / d + h * (MP2_x - MP1_x) / d + MP1_y;
        }
    }
}
