using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
//using System.Web.UI.WebControls;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Draw;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.Win32;
using System.Data;
using System.IO;
using System.Web.Configuration;
using DataBaseWrapper;
using System.Globalization;


namespace Excel_Hochladen_und_Einlesen
{
    public partial class Rechnung_erstellen : System.Web.UI.Page
    {
        string connStrg = WebConfigurationManager.ConnectionStrings["AppDbInt"].ConnectionString;
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btn_RechnungErstellen_Click(object sender, EventArgs e)
        {
            Rechnung("50000", 3, 22);
            Gutschrift("50664", 4, 22);
            RechnungMonatlich("50000", 10, 22);
            GutschriftMonatlich("50664", 10, 22);


        }

        public void GutschriftMonatlich(string mitgliedID, int monat, int year)
        {



            //MemoryStream ms = new MemoryStream();

            //PdfWriter w = new PdfWriter(ms);
            //PdfDocument pdf = new PdfDocument(w);
            //Document doc = new Document(pdf);

            PdfWriter writer = new PdfWriter("Z:\\5. Klasse\\Energiegemeinschaften Projekt\\Excel_Hochladen_und_Einlesen\\Excel_Hochladen_und_Einlesen\\bin\\GutschriftMonatlich.pdf");
            PdfDocument pdf = new PdfDocument(writer);
            Document doc = new Document(pdf);


            string connStrg = WebConfigurationManager.ConnectionStrings["AppDbInt"].ConnectionString;
            DataBase12 database = new DataBase12(connStrg);


            //Adresse?? Obamnn?? über die Gemein ID


            string monatS = "";
   

            string verbrauch1 = "0";
     


            //Get metering point
            string sqlCmd = $"SELECT MeteringpointID FROM eg_mitglieder where MitgliedsID = '{mitgliedID}'";
            string meteringpoint = Convert.ToString(database.RunQueryScalar(sqlCmd));

            if (monat == 1) monatS = "Januar";
            if (monat == 2) monatS = "Februar";
            if (monat == 3) monatS = "März";
            if (monat == 4) monatS = "April";
            if (monat == 5) monatS = "Mail";
            if (monat == 6) monatS = "Juni";
            if (monat == 7) monatS = "Juli";
            if (monat == 8) monatS = "August";
            if (monat == 9) monatS = "September";
            if (monat == 10) monatS = "Oktober";
            if (monat == 11) monatS = "November";
            if (monat == 12) monatS = "Dezember";


            //Erzeugung
            sqlCmd = $"SELECT Erzeugung_in_kWh, DatumVon FROM eg_monatlicheerzeugung where Mitglied = '{mitgliedID}'";
            DataTable dtVerbrauch = database.RunQuery(sqlCmd);




            foreach (DataRow dr in dtVerbrauch.Rows)
            {
                DateTime date = DateTime.Parse(dr[1].ToString());
                if (date.Month == monat && date.Year == (year+2000)) verbrauch1 = dr[0].ToString();    
            }

            decimal eGes = Convert.ToDecimal(verbrauch1);


            //Quote - Einspeistarif ist gleich gesamte EEG
            sqlCmd = $"SELECT eg_raten.Wert FROM eg_raten Where Bezeichnung = 'Einspeiser'";
            decimal quote = Convert.ToDecimal(Math.Round(Convert.ToDouble(database.RunQueryScalar(sqlCmd)), 2));

            //Kosten
            decimal preis = quote * eGes;
            preis = Math.Round(preis, 2);

            // EG Name
            sqlCmd = $"SELECT eg_energiegemeinschaft.Name FROM eg_Mitglieder LEFT JOIN eg_energiegemeinschaft ON  eg_energiegemeinschaft.GemID = eg_Mitglieder.GemeinID WHERE eg_Mitglieder.MitgliedsID = {mitgliedID}";
            string EGName = Convert.ToString(database.RunQueryScalar(sqlCmd));

            // Adresse Mitglied und name
            sqlCmd = $"SELECT Adresse, Postleitzahl, Ort, Vorname, Nachname FROM eg_mitglieder Where MitgliedsID = {mitgliedID}";
            DataTable dtAddress = database.RunQuery(sqlCmd);

            //Adresse EGG
            sqlCmd = $"SELECT eg_mitglieder.Adresse, eg_mitglieder.Postleitzahl, eg_mitglieder.Ort FROM eg_energiegemeinschaft LEFT JOIN eg_mitglieder ON eg_energiegemeinschaft.GruenderID = eg_mitglieder.MitgliedsID Where eg_energiegemeinschaft.Name = '{EGName}'";
            DataTable dtAddGruender = database.RunQuery(sqlCmd);

            //IBAN und BIC - vom Erzeuger
            sqlCmd = $"SELECT IBAN, BIC FROM eg_mitglieder Where MitgliedsID = {mitgliedID}";
            DataTable dtIBAN = database.RunQuery(sqlCmd);

            //ZVR-Zahl
            sqlCmd = $"SELECT ZVR FROM eg_energiegemeinschaft WHERE Name = '{EGName}'";
            string zVRZahl = database.RunQueryScalar(sqlCmd).ToString();




            doc.Add(new Paragraph($@"Energiegemeinschaft {EGName}").AddTabStops(new TabStop(700, TabAlignment.RIGHT), new TabStop(0, TabAlignment.LEFT)).Add(new Tab()).Add(@"ZVR-Zahl
").Add(dtAddGruender.Rows[0].ItemArray[0].ToString()).Add(new Tab()).Add($@"{zVRZahl}
").Add($"{dtAddGruender.Rows[0].ItemArray[1].ToString()} {dtAddGruender.Rows[0].ItemArray[2].ToString()}"));






            LineSeparator ls = new LineSeparator(new SolidLine());
            doc.Add(ls);

            Paragraph header = new Paragraph("Gutschrift")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBold()
                .SetFontSize(50);
            doc.Add(header);

            doc.Add(new Paragraph());

            Paragraph subheader = new Paragraph($@"Stromlieferung {monatS}/20{year}
    Energiegemeinschaft {EGName}")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBold()
           .SetFontSize(25);
            doc.Add(subheader);



            doc.Add(ls);

            doc.Add(new Paragraph());

            Paragraph Sub = new Paragraph("1  Erzeugungsanlage")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Sub);

            Paragraph text = new Paragraph($"Zählpunkt:         {meteringpoint}  ")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetFontSize(15);
            doc.Add(text);
            Paragraph Anlagestandort = new Paragraph($"Anlagestandort:    {dtAddress.Rows[0].ItemArray[0].ToString()}, {dtAddress.Rows[0].ItemArray[1].ToString()} {dtAddress.Rows[0].ItemArray[2].ToString()}")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetFontSize(15);
            doc.Add(Anlagestandort);

            Paragraph Verbrauch = new Paragraph("2  Produktion (lt. EDA-Portal)")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Verbrauch);


            Table table = new Table(2, true);

            Cell cell11 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("Monat"));
            Cell cell15 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBold()
               .Add(new Paragraph(monatS).SetBold());

            Cell cell21 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("Produktion [kWh]"));
            Cell cell25 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBold()
               .Add(new Paragraph(eGes.ToString()));



            table.AddCell(cell11);
            table.AddCell(cell15);
            table.AddCell(cell21);
            table.AddCell(cell25);

            doc.Add(table);

            Paragraph Tarif = new Paragraph("3  Tarif (lt. Generalversammlung)")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Tarif);

            Paragraph Zählpunkt = new Paragraph($"Einspeisepreis:  {quote} € / kWh")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetBold()
                .SetFontSize(15);
            doc.Add(Zählpunkt);


            Paragraph Energiekosten = new Paragraph("4  Energievergütung")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Energiekosten);



            //  Paragraph Kosten = new Paragraph($"Q{quartal}/20{year}:            {gesKost}")
            //.SetTextAlignment(TextAlignment.LEFT)
            //.SetBold()
            //.SetFontSize(30);
            //  doc.Add(Kosten);

            Table table2 = new Table(1, true);

            Cell cell112 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph($"{monatS}/20{year}:                      €{preis.ToString()}")
          .SetTextAlignment(TextAlignment.LEFT)
          .SetBold()
          .SetFontSize(32)
            );

            table2.AddCell(cell112);

            doc.Add(table2);

            Text Boldtext = new Text($"{dtAddress.Rows[0].ItemArray[3]} {dtAddress.Rows[0].ItemArray[4]}, IBAN {dtIBAN.Rows[0].ItemArray[0]}, BIC {dtIBAN.Rows[0].ItemArray[1]}").SetBold();
            Paragraph blabla = new Paragraph(@"Überweisung der Gutschrift erfolgt auf folgendes Konto:
").Add(Boldtext)
               .SetTextAlignment(TextAlignment.LEFT)
               .SetFontSize(12);
            doc.Add(blabla);

            doc.Add(new Paragraph());
            doc.Add(new Paragraph());
            doc.Add(new Paragraph());
            doc.Add(new Paragraph());

            doc.Add(new Paragraph(@"________________________").AddTabStops(new TabStop(700, TabAlignment.RIGHT), new TabStop(0, TabAlignment.LEFT)).Add(new Tab()).Add(@"________________________
")
                .Add("(Obmann)").Add(new Tab()).Add(@"(Kassier)"));

            doc.Add(new Paragraph());
            doc.Add(new Paragraph());
            doc.Add(new Paragraph());

            doc.Add(new Paragraph($@"Gutschrift Lieferung {monatS} / 20{year} ,").AddTabStops(new TabStop(700, TabAlignment.RIGHT), new TabStop(0, TabAlignment.LEFT)).Add(new Tab()).Add($@"{meteringpoint}"));



            doc.Close();
            //byte[] bytesInStream = ms.ToArray();
            //return bytesInStream;
        }

        public void RechnungMonatlich(string mitgliedID, int monat, int year)
        {
            //year = year - 2000;

            //MemoryStream ms = new MemoryStream();

            //PdfWriter w = new PdfWriter(ms);
            //PdfWriter writer = new PdfWriter("Z:\\5. Klasse\\Energiegemeinschaften Projekt\\Excel_Hochladen_und_Einlesen\\Excel_Hochladen_und_Einlesen\\bin\\RechnungMonatlich.pdf");
            //PdfDocument pdf = new PdfDocument(w);
            //Document doc = new Document(pdf);

            PdfWriter writer = new PdfWriter("Z:\\5. Klasse\\Energiegemeinschaften Projekt\\Excel_Hochladen_und_Einlesen\\Excel_Hochladen_und_Einlesen\\bin\\RechnungMonatlich.pdf");
            PdfDocument pdf = new PdfDocument(writer);
            Document doc = new Document(pdf);


            string connStrg = WebConfigurationManager.ConnectionStrings["AppDbInt"].ConnectionString;
            DataBase12 database = new DataBase12(connStrg);


            //Adresse?? Obamnn?? über die Gemein ID

            string monatS = "";
            string verbrauch1 = "0";
      

            //Get metering point
            string sqlCmd = $"SELECT MeteringpointID FROM eg_mitglieder where MitgliedsID = '{mitgliedID}'";
            string meteringpoint = Convert.ToString(database.RunQueryScalar(sqlCmd));

            if (monat == 1) monatS = "Januar";
            if (monat == 2) monatS = "Februar";
            if (monat == 3) monatS = "März";
            if (monat == 4) monatS = "April";
            if (monat == 5) monatS = "Mail";
            if (monat == 6) monatS = "Juni";
            if (monat == 7) monatS = "Juli";
            if (monat == 8) monatS = "August";
            if (monat == 9) monatS = "September";
            if (monat == 10) monatS = "Oktober";
            if (monat == 11) monatS = "November";
            if (monat == 12) monatS = "Dezember";



            //Verbrauch
            sqlCmd = $"SELECT Verbrauch_in_kWh, DatumVon FROM eg_monatlicherverbrauch where Mitglied = '{mitgliedID}'";
            DataTable dtVerbrauch = database.RunQuery(sqlCmd);

            foreach (DataRow dr in dtVerbrauch.Rows)

            {
                DateTime date = DateTime.Parse(dr[1].ToString());
                            
                    if (date.Month == monat && date.Year == (year+2000)) verbrauch1 = dr[0].ToString();
           
            }

            decimal vGes = Convert.ToDecimal(verbrauch1);


            //Quote
            sqlCmd = $"SELECT eg_raten.Wert FROM eg_Mitglieder LEFT JOIN  eg_raten ON eg_mitglieder.RatenBezeichnung = eg_raten.Bezeichnung Where MitgliedsID = {mitgliedID}";
            decimal quote = Convert.ToDecimal(Math.Round(Convert.ToDouble(database.RunQueryScalar(sqlCmd)), 2));

            //Kosten
            decimal preis = quote * vGes;
            preis = Math.Round(preis, 2);

            // EG Name
            sqlCmd = $"SELECT eg_energiegemeinschaft.Name FROM eg_Mitglieder LEFT JOIN eg_energiegemeinschaft ON  eg_energiegemeinschaft.GemID = eg_Mitglieder.GemeinID WHERE eg_Mitglieder.MitgliedsID = {mitgliedID}";
            string EGName = Convert.ToString(database.RunQueryScalar(sqlCmd));

            // Adresse Mitglied
            sqlCmd = $"SELECT Adresse, Postleitzahl, Ort FROM eg_mitglieder Where MitgliedsID = {mitgliedID}";
            DataTable dtAddress = database.RunQuery(sqlCmd);

            //Adresse EGG
            sqlCmd = $"SELECT eg_mitglieder.Adresse, eg_mitglieder.Postleitzahl, eg_mitglieder.Ort FROM eg_energiegemeinschaft LEFT JOIN eg_mitglieder ON eg_energiegemeinschaft.GruenderID = eg_mitglieder.MitgliedsID Where eg_energiegemeinschaft.Name = '{EGName}'";
            DataTable dtAddGruender = database.RunQuery(sqlCmd);

            //IBAN und BIC
            sqlCmd = $"SELECT IBAN, BIC, ZVR FROM eg_energiegemeinschaft WHERE Name = '{EGName}'";
            DataTable dtIBAN = database.RunQuery(sqlCmd);


            string zVRZahl = dtIBAN.Rows[0][2].ToString();



            doc.Add(new Paragraph($@"Energiegemeinschaft {EGName}").AddTabStops(new TabStop(700, TabAlignment.RIGHT), new TabStop(0, TabAlignment.LEFT)).Add(new Tab()).Add(@"ZVR-Zahl
").Add(dtAddGruender.Rows[0].ItemArray[0].ToString()).Add(new Tab()).Add($@"{zVRZahl}
").Add($"{dtAddGruender.Rows[0].ItemArray[1].ToString()} {dtAddGruender.Rows[0].ItemArray[2].ToString()}"));






            LineSeparator ls = new LineSeparator(new SolidLine());
            doc.Add(ls);

            Paragraph header = new Paragraph("Rechnung")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBold()
                .SetFontSize(50);
            doc.Add(header);

            doc.Add(new Paragraph());

            Paragraph subheader = new Paragraph($@"Strombezug {monatS}/20{year}
    Energiegemeinschaft {EGName}")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBold()
           .SetFontSize(25);
            doc.Add(subheader);



            doc.Add(ls);

            doc.Add(new Paragraph());

            Paragraph Sub = new Paragraph("1  Verbrauchsanlage")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Sub);

            Paragraph text = new Paragraph($"Zählpunkt:         {meteringpoint}  ")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetFontSize(15);
            doc.Add(text);
            Paragraph Anlagestandort = new Paragraph($"Anlagestandort:    {dtAddress.Rows[0].ItemArray[0].ToString()}, {dtAddress.Rows[0].ItemArray[1].ToString()} {dtAddress.Rows[0].ItemArray[2].ToString()}")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetFontSize(15);
            doc.Add(Anlagestandort);

            Paragraph Verbrauch = new Paragraph("2  Verbrauch (lt. EDA-Portal)")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Verbrauch);


            Table table = new Table(2, true);

            Cell cell11 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("Monat"));
        
            Cell cell15 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBold()
               .Add(new Paragraph($"{monatS}").SetBold());

            Cell cell21 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("Verbrauch [kWh]"));
            Cell cell25 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBold()
               .Add(new Paragraph(vGes.ToString()));



            table.AddCell(cell11);
            table.AddCell(cell15);
            table.AddCell(cell21);
            table.AddCell(cell25);

            doc.Add(table);

            Paragraph Tarif = new Paragraph("3  Tarif (lt. Generalversammlung)")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Tarif);

            Paragraph Zählpunkt = new Paragraph($"Arbeitspreis:  {quote} € / kWh")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetBold()
                .SetFontSize(15);
            doc.Add(Zählpunkt);


            Paragraph Energiekosten = new Paragraph("4  Energiekosten")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Energiekosten);



            //  Paragraph Kosten = new Paragraph($"Q{quartal}/20{year}:            {gesKost}")
            //.SetTextAlignment(TextAlignment.LEFT)
            //.SetBold()
            //.SetFontSize(30);
            //  doc.Add(Kosten);

            Table table2 = new Table(1, true);

            Cell cell112 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph($"{monatS}/20{year}:                      €{preis.ToString()}")
          .SetTextAlignment(TextAlignment.LEFT)
          .SetBold()
          .SetFontSize(32)
            );

            table2.AddCell(cell112);

            doc.Add(table2);

            Text Boldtext = new Text($"EEG {EGName}, IBAN {dtIBAN.Rows[0].ItemArray[0]}, BIC {dtIBAN.Rows[0].ItemArray[1]}").SetBold();
            Paragraph blabla = new Paragraph(@"Steuerbefreit – Kleinunternehmer gemäß § 6 Abs. 1 / 27 UStG
zahlbar ohne Abzüge innerhalb der nächsten 30 Tage auf folgendes Konto:
").Add(Boldtext)
               .SetTextAlignment(TextAlignment.LEFT)
               .SetFontSize(12);
            doc.Add(blabla);

            doc.Add(new Paragraph());
            doc.Add(new Paragraph());
            doc.Add(new Paragraph());
            doc.Add(new Paragraph());

            doc.Add(new Paragraph(@"________________________").AddTabStops(new TabStop(700, TabAlignment.RIGHT), new TabStop(0, TabAlignment.LEFT)).Add(new Tab()).Add(@"________________________
")
                .Add("(Obmann)").Add(new Tab()).Add(@"(Kassier)"));

            doc.Add(new Paragraph());
            doc.Add(new Paragraph());
            doc.Add(new Paragraph());

            doc.Add(new Paragraph($@"Rechnung Bezug {monatS} / 20{year} ,").AddTabStops(new TabStop(700, TabAlignment.RIGHT), new TabStop(0, TabAlignment.LEFT)).Add(new Tab()).Add($@"{meteringpoint}"));



            doc.Close();

            //byte[] bytesInStream = ms.ToArray();
           // return bytesInStream;







        }




        public void Gutschrift(string mitgliedID, int quartal, int year)
        {



            PdfWriter writer = new PdfWriter("Z:\\5. Klasse\\Energiegemeinschaften Projekt\\Excel_Hochladen_und_Einlesen\\Excel_Hochladen_und_Einlesen\\bin\\Gutschrift.pdf");
            PdfDocument pdf = new PdfDocument(writer);
            Document doc = new Document(pdf);


            string connStrg = WebConfigurationManager.ConnectionStrings["AppDbInt"].ConnectionString;
            DataBase12 database = new DataBase12(connStrg);


            //Adresse?? Obamnn?? über die Gemein ID


            string monat1 = "";
            string monat2 = "";
            string monat3 = "";

            string verbrauch1 = "0";
            string verbrauch2 = "0";
            string verbrauch3 = "0";


            //Get metering point
            string sqlCmd = $"SELECT MeteringpointID FROM eg_mitglieder where MitgliedsID = '{mitgliedID}'";
            string meteringpoint = Convert.ToString(database.RunQueryScalar(sqlCmd));

            if (quartal == 1)
            {
                monat1 = "Januar";
                monat2 = "Februar";
                monat3 = "März";
            }
            if (quartal == 2)
            {
                monat1 = "April";
                monat2 = "Mai";
                monat3 = "Juni";
            }
            if (quartal == 3)
            {
                monat1 = "Juli";
                monat2 = "August";
                monat3 = "September";
            }
            if (quartal == 4)
            {
                monat1 = "Oktober";
                monat2 = "November";
                monat3 = "Dezember";
            }


            //Erzeugung
            sqlCmd = $"SELECT Erzeugung_in_kWh, DatumVon FROM eg_monatlicheerzeugung where Mitglied = '{mitgliedID}'";
            DataTable dtVerbrauch = database.RunQuery(sqlCmd);

            foreach (DataRow dr in dtVerbrauch.Rows)

            {
                string date = dr[1].ToString();
                if (date == $"01.01.20{year} 00:00:00") verbrauch1 = dr[0].ToString();
                if (date == $"01.02.20{year} 00:00:00") verbrauch2 = dr[0].ToString();
                if (date == $"01.03.20{year} 00:00:00") verbrauch3 = dr[0].ToString();

                if (date == $"01.04.20{year} 00:00:00") verbrauch1 = dr[0].ToString();
                if (date == $"01.05.20{year} 00:00:00") verbrauch2 = dr[0].ToString();
                if (date == $"01.06.20{year} 00:00:00") verbrauch3 = dr[0].ToString();

                if (date == $"01.07.20{year} 00:00:00") verbrauch1 = dr[0].ToString();
                if (date == $"01.08.20{year} 00:00:00") verbrauch2 = dr[0].ToString();
                if (date == $"01.09.20{year} 00:00:00") verbrauch3 = dr[0].ToString();

                if (date == $"01.10.20{year} 00:00:00") verbrauch1 = dr[0].ToString();
                if (date == $"01.11.20{year} 00:00:00") verbrauch2 = dr[0].ToString();
                if (date == $"01.12.20{year} 00:00:00") verbrauch3 = dr[0].ToString();
            }

            decimal eGes = Convert.ToDecimal(verbrauch1) + Convert.ToDecimal(verbrauch2) + Convert.ToDecimal(verbrauch3);


            //Quote - Einspeistarif ist gleich gesamte EEG
            sqlCmd = $"SELECT eg_raten.Wert FROM eg_raten Where Bezeichnung = 'Einspeiser'";
            decimal quote = Convert.ToDecimal(Math.Round(Convert.ToDouble(database.RunQueryScalar(sqlCmd)), 2));

            //Kosten
            decimal preis = quote * eGes;
            preis = Math.Round(preis, 2);

            // EG Name
            sqlCmd = $"SELECT eg_energiegemeinschaft.Name FROM eg_Mitglieder LEFT JOIN eg_energiegemeinschaft ON  eg_energiegemeinschaft.GemID = eg_Mitglieder.GemeinID WHERE eg_Mitglieder.MitgliedsID = {mitgliedID}";
            string EGName = Convert.ToString(database.RunQueryScalar(sqlCmd));

            // Adresse Mitglied und name
            sqlCmd = $"SELECT Adresse, Postleitzahl, Ort, Vorname, Nachname FROM eg_mitglieder Where MitgliedsID = {mitgliedID}";
            DataTable dtAddress = database.RunQuery(sqlCmd);

            //Adresse EGG
            sqlCmd = $"SELECT eg_mitglieder.Adresse, eg_mitglieder.Postleitzahl, eg_mitglieder.Ort FROM eg_energiegemeinschaft LEFT JOIN eg_mitglieder ON eg_energiegemeinschaft.GruenderID = eg_mitglieder.MitgliedsID Where eg_energiegemeinschaft.Name = '{EGName}'";
            DataTable dtAddGruender = database.RunQuery(sqlCmd);

            //IBAN und BIC - vom Erzeuger
            sqlCmd = $"SELECT IBAN, BIC FROM eg_mitglieder Where MitgliedsID = {mitgliedID}";
            DataTable dtIBAN = database.RunQuery(sqlCmd);

            //ZVR-Zahl
            sqlCmd = $"SELECT ZVR FROM eg_energiegemeinschaft WHERE Name = '{EGName}'";
            string zVRZahl = database.RunQueryScalar(sqlCmd).ToString();
             



            doc.Add(new Paragraph($@"Energiegemeinschaft {EGName}").AddTabStops(new TabStop(700, TabAlignment.RIGHT), new TabStop(0, TabAlignment.LEFT)).Add(new Tab()).Add(@"ZVR-Zahl
").Add(dtAddGruender.Rows[0].ItemArray[0].ToString()).Add(new Tab()).Add($@"{zVRZahl}
").Add($"{dtAddGruender.Rows[0].ItemArray[1].ToString()} {dtAddGruender.Rows[0].ItemArray[2].ToString()}"));






            LineSeparator ls = new LineSeparator(new SolidLine());
            doc.Add(ls);

            Paragraph header = new Paragraph("Gutschrift")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBold()
                .SetFontSize(50);
            doc.Add(header);

            doc.Add(new Paragraph());

            Paragraph subheader = new Paragraph($@"Stromlieferung Q{quartal}/20{year}
    Energiegemeinschaft {EGName}")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBold()
           .SetFontSize(25);
            doc.Add(subheader);



            doc.Add(ls);

            doc.Add(new Paragraph());

            Paragraph Sub = new Paragraph("1  Erzeugungsanlage")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Sub);

            Paragraph text = new Paragraph($"Zählpunkt:         {meteringpoint}  ")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetFontSize(15);
            doc.Add(text);
            Paragraph Anlagestandort = new Paragraph($"Anlagestandort:    {dtAddress.Rows[0].ItemArray[0].ToString()}, {dtAddress.Rows[0].ItemArray[1].ToString()} {dtAddress.Rows[0].ItemArray[2].ToString()}")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetFontSize(15);
            doc.Add(Anlagestandort);

            Paragraph Verbrauch = new Paragraph("2  Produktion (lt. EDA-Portal)")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Verbrauch);


            Table table = new Table(5, true);

            Cell cell11 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("Monat"));
            Cell cell12 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph(monat1));
            Cell cell13 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph(monat2));
            Cell cell14 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph(monat3));
            Cell cell15 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBold()
               .Add(new Paragraph("Gesamt").SetBold());

            Cell cell21 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("Produktion [kWh]"));
            Cell cell22 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph(verbrauch1));
            Cell cell23 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph(verbrauch2));
            Cell cell24 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph(verbrauch3));
            Cell cell25 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBold()
               .Add(new Paragraph(eGes.ToString()));



            table.AddCell(cell11);
            table.AddCell(cell12);
            table.AddCell(cell13);
            table.AddCell(cell14);
            table.AddCell(cell15);
            table.AddCell(cell21);
            table.AddCell(cell22);
            table.AddCell(cell23);
            table.AddCell(cell24);
            table.AddCell(cell25);

            doc.Add(table);

            Paragraph Tarif = new Paragraph("3  Tarif (lt. Generalversammlung)")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Tarif);

            Paragraph Zählpunkt = new Paragraph($"Einspeisepreis:  {quote} € / kWh")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetBold()
                .SetFontSize(15);
            doc.Add(Zählpunkt);


            Paragraph Energiekosten = new Paragraph("4  Energievergütung")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Energiekosten);



            //  Paragraph Kosten = new Paragraph($"Q{quartal}/20{year}:            {gesKost}")
            //.SetTextAlignment(TextAlignment.LEFT)
            //.SetBold()
            //.SetFontSize(30);
            //  doc.Add(Kosten);

            Table table2 = new Table(1, true);

            Cell cell112 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph($"Q{quartal}/20{year}:                      €{preis.ToString()}")
          .SetTextAlignment(TextAlignment.LEFT)
          .SetBold()
          .SetFontSize(32)
            );

            table2.AddCell(cell112);

            doc.Add(table2);

            Text Boldtext = new Text($"{dtAddress.Rows[0].ItemArray[3]} {dtAddress.Rows[0].ItemArray[4]}, IBAN {dtIBAN.Rows[0].ItemArray[0]}, BIC {dtIBAN.Rows[0].ItemArray[1]}").SetBold();
            Paragraph blabla = new Paragraph(@"Überweisung der Gutschrift erfolgt auf folgendes Konto:
").Add(Boldtext)
               .SetTextAlignment(TextAlignment.LEFT)
               .SetFontSize(12);
            doc.Add(blabla);

            doc.Add(new Paragraph());
            doc.Add(new Paragraph());
            doc.Add(new Paragraph());
            doc.Add(new Paragraph());

            doc.Add(new Paragraph(@"________________________").AddTabStops(new TabStop(700, TabAlignment.RIGHT), new TabStop(0, TabAlignment.LEFT)).Add(new Tab()).Add(@"________________________
")
                .Add("(Obmann)").Add(new Tab()).Add(@"(Kassier)"));

            doc.Add(new Paragraph());
            doc.Add(new Paragraph());
            doc.Add(new Paragraph());

            doc.Add(new Paragraph($@"Gutschrift Lieferung Q{quartal} / 20{year} ,").AddTabStops(new TabStop(700, TabAlignment.RIGHT), new TabStop(0, TabAlignment.LEFT)).Add(new Tab()).Add($@"{meteringpoint}"));



            doc.Close();

        }



        public void Rechnung(string mitgliedID, int quartal, int year)  
        {
            PdfWriter writer = new PdfWriter("Z:\\5. Klasse\\Energiegemeinschaften Projekt\\Excel_Hochladen_und_Einlesen\\Excel_Hochladen_und_Einlesen\\bin\\Rechnung.pdf");
            PdfDocument pdf = new PdfDocument(writer);
            Document doc = new Document(pdf);


            string connStrg = WebConfigurationManager.ConnectionStrings["AppDbInt"].ConnectionString;
            DataBase12 database = new DataBase12(connStrg);


            //Adresse?? Obamnn?? über die Gemein ID


            string monat1 = "";
            string monat2 = "";
            string monat3 = "";

            string verbrauch1 = "0";
            string verbrauch2 = "0";
            string verbrauch3 = "0";


            //Get metering point
            string sqlCmd = $"SELECT MeteringpointID FROM eg_mitglieder where MitgliedsID = '{mitgliedID}'";
            string meteringpoint = Convert.ToString(database.RunQueryScalar(sqlCmd));

            if (quartal == 1)
            {
                monat1 = "Januar";
                monat2 = "Februar";
                monat3 = "März";
            }
            if (quartal == 2)
            {
                monat1 = "April";
                monat2 = "Mai";
                monat3 = "Juni";
            }
            if (quartal == 3)
            {
                monat1 = "Juli";
                monat2 = "August";
                monat3 = "September";
            }
            if (quartal == 4)
            {
                monat1 = "Oktober";
                monat2 = "November";
                monat3 = "Dezember";
            }


            //Verbrauch
            sqlCmd = $"SELECT Verbrauch_in_kWh, DatumVon FROM eg_monatlicherverbrauch where Mitglied = '{mitgliedID}'";
            DataTable dtVerbrauch = database.RunQuery(sqlCmd);

            foreach (DataRow dr in dtVerbrauch.Rows)

            {
                string date = dr[1].ToString();
                if (date == $"01.01.20{year} 00:00:00") verbrauch1 = dr[0].ToString();
                if (date == $"01.02.20{year} 00:00:00") verbrauch2 = dr[0].ToString();
                if (date == $"01.03.20{year} 00:00:00") verbrauch3 = dr[0].ToString();

                if (date == $"01.04.20{year} 00:00:00") verbrauch1 = dr[0].ToString();
                if (date == $"01.05.20{year} 00:00:00") verbrauch2 = dr[0].ToString();
                if (date == $"01.06.20{year} 00:00:00") verbrauch3 = dr[0].ToString();

                if (date == $"01.07.20{year} 00:00:00") verbrauch1 = dr[0].ToString();
                if (date == $"01.08.20{year} 00:00:00") verbrauch2 = dr[0].ToString();
                if (date == $"01.09.20{year} 00:00:00") verbrauch3 = dr[0].ToString();

                if (date == $"01.10.20{year} 00:00:00") verbrauch1 = dr[0].ToString();
                if (date == $"01.11.20{year} 00:00:00") verbrauch2 = dr[0].ToString();
                if (date == $"01.12.20{year} 00:00:00") verbrauch3 = dr[0].ToString();
            }

            decimal vGes = Convert.ToDecimal(verbrauch1) + Convert.ToDecimal(verbrauch2) + Convert.ToDecimal(verbrauch3);


            //Quote
            sqlCmd = $"SELECT eg_raten.Wert FROM eg_Mitglieder LEFT JOIN  eg_raten ON eg_mitglieder.RatenBezeichnung = eg_raten.Bezeichnung Where MitgliedsID = {mitgliedID}";
            decimal quote = Convert.ToDecimal(Math.Round(Convert.ToDouble(database.RunQueryScalar(sqlCmd)), 2));

            //Kosten
            decimal preis = quote * vGes;
            preis = Math.Round(preis, 2);

            // EG Name
            sqlCmd = $"SELECT eg_energiegemeinschaft.Name FROM eg_Mitglieder LEFT JOIN eg_energiegemeinschaft ON  eg_energiegemeinschaft.GemID = eg_Mitglieder.GemeinID WHERE eg_Mitglieder.MitgliedsID = {mitgliedID}";
            string EGName = Convert.ToString(database.RunQueryScalar(sqlCmd));

            // Adresse Mitglied
            sqlCmd = $"SELECT Adresse, Postleitzahl, Ort FROM eg_mitglieder Where MitgliedsID = {mitgliedID}";
            DataTable dtAddress = database.RunQuery(sqlCmd);

            //Adresse EGG
            sqlCmd = $"SELECT eg_mitglieder.Adresse, eg_mitglieder.Postleitzahl, eg_mitglieder.Ort FROM eg_energiegemeinschaft LEFT JOIN eg_mitglieder ON eg_energiegemeinschaft.GruenderID = eg_mitglieder.MitgliedsID Where eg_energiegemeinschaft.Name = '{EGName}'";
            DataTable dtAddGruender = database.RunQuery(sqlCmd);

            //IBAN und BIC
            sqlCmd = $"SELECT IBAN, BIC FROM eg_energiegemeinschaft WHERE Name = '{EGName}'";
            DataTable dtIBAN = database.RunQuery(sqlCmd);


            //ZVR-Zahl
            sqlCmd = $"SELECT ZVR FROM eg_energiegemeinschaft WHERE Name = '{EGName}'";
            string zVRZahl = database.RunQueryScalar(sqlCmd).ToString();


            doc.Add(new Paragraph($@"Energiegemeinschaft {EGName}").AddTabStops(new TabStop(700, TabAlignment.RIGHT), new TabStop(0, TabAlignment.LEFT)).Add(new Tab()).Add(@"ZVR-Zahl
").Add(dtAddGruender.Rows[0].ItemArray[0].ToString()).Add(new Tab()).Add($@"{zVRZahl}
").Add($"{dtAddGruender.Rows[0].ItemArray[1].ToString()} {dtAddGruender.Rows[0].ItemArray[2].ToString()}"));






            LineSeparator ls = new LineSeparator(new SolidLine());
            doc.Add(ls);

            Paragraph header = new Paragraph("Rechnung")
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBold()
                .SetFontSize(50);
            doc.Add(header);

            doc.Add(new Paragraph());

            Paragraph subheader = new Paragraph($@"Strombezug Q{quartal}/20{year}
    Energiegemeinschaft {EGName}")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBold()
           .SetFontSize(25);
            doc.Add(subheader);



            doc.Add(ls);

            doc.Add(new Paragraph());

            Paragraph Sub = new Paragraph("1  Verbrauchsanlage")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Sub);

            Paragraph text = new Paragraph($"Zählpunkt:         {meteringpoint}  ")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetFontSize(15);
            doc.Add(text);
            Paragraph Anlagestandort = new Paragraph($"Anlagestandort:    {dtAddress.Rows[0].ItemArray[0].ToString()}, {dtAddress.Rows[0].ItemArray[1].ToString()} {dtAddress.Rows[0].ItemArray[2].ToString()}")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetFontSize(15);
            doc.Add(Anlagestandort);

            Paragraph Verbrauch = new Paragraph("2  Verbrauch (lt. EDA-Portal)")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Verbrauch);


            Table table = new Table(5, true);

            Cell cell11 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("Monat"));
            Cell cell12 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph(monat1));
            Cell cell13 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph(monat2));
            Cell cell14 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph(monat3));
            Cell cell15 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBold()
               .Add(new Paragraph("Gesamt").SetBold());

            Cell cell21 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("Verbrauch [kWh]"));
            Cell cell22 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph(verbrauch1));
            Cell cell23 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph(verbrauch2));
            Cell cell24 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph(verbrauch3));
            Cell cell25 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBold()
               .Add(new Paragraph(vGes.ToString()));



            table.AddCell(cell11);
            table.AddCell(cell12);
            table.AddCell(cell13);
            table.AddCell(cell14);
            table.AddCell(cell15);
            table.AddCell(cell21);
            table.AddCell(cell22);
            table.AddCell(cell23);
            table.AddCell(cell24);
            table.AddCell(cell25);

            doc.Add(table);

            Paragraph Tarif = new Paragraph("3  Tarif (lt. Generalversammlung)")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Tarif);

            Paragraph Zählpunkt = new Paragraph($"Arbeitspreis:  {quote} Euro / kWh")
                .SetTextAlignment(TextAlignment.LEFT)
                .SetBold()
                .SetFontSize(15);
            doc.Add(Zählpunkt);


            Paragraph Energiekosten = new Paragraph("4  Energiekosten")
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBold()
               .SetFontSize(15);
            doc.Add(Energiekosten);



            //  Paragraph Kosten = new Paragraph($"Q{quartal}/20{year}:            {gesKost}")
            //.SetTextAlignment(TextAlignment.LEFT)
            //.SetBold()
            //.SetFontSize(30);
            //  doc.Add(Kosten);

            Table table2 = new Table(1, true);

            Cell cell112 = new Cell(1, 1)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph($"Q{quartal}/20{year}:                      €{preis.ToString()}")
          .SetTextAlignment(TextAlignment.LEFT)
          .SetBold()
          .SetFontSize(32)
            );

            table2.AddCell(cell112);

            doc.Add(table2);

            Text Boldtext = new Text($"EEG {EGName}, IBAN {dtIBAN.Rows[0].ItemArray[0]}, BIC {dtIBAN.Rows[0].ItemArray[1]}").SetBold();
            Paragraph blabla = new Paragraph(@"Steuerbefreit – Kleinunternehmer gemäß § 6 Abs. 1 / 27 UStG
zahlbar ohne Abzüge innerhalb der nächsten 30 Tage auf folgendes Konto:
").Add(Boldtext)
               .SetTextAlignment(TextAlignment.LEFT)
               .SetFontSize(12);
            doc.Add(blabla);

            doc.Add(new Paragraph());
            doc.Add(new Paragraph());
            doc.Add(new Paragraph());
            doc.Add(new Paragraph());

            doc.Add(new Paragraph(@"________________________").AddTabStops(new TabStop(700, TabAlignment.RIGHT), new TabStop(0, TabAlignment.LEFT)).Add(new Tab()).Add(@"________________________
")
                .Add("(Obmann)").Add(new Tab()).Add(@"(Kassier)"));

            doc.Add(new Paragraph());
            doc.Add(new Paragraph());
            doc.Add(new Paragraph());

            doc.Add(new Paragraph($@"Rechnung Bezug Q{quartal} / 20{year} ,").AddTabStops(new TabStop(700, TabAlignment.RIGHT), new TabStop(0, TabAlignment.LEFT)).Add(new Tab()).Add($@"{meteringpoint}"));



            doc.Close();

        }
    }
}
