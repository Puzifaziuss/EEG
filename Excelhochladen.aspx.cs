using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Configuration;
using System.Web.UI;
using System.Web.UI.WebControls;
using DataBaseWrapper;
using System.Globalization;

namespace Excel_Hochladen_und_Einlesen
{
    public partial class Excelhochladen : System.Web.UI.Page
    {
        string connStrg = WebConfigurationManager.ConnectionStrings["AppDbInt"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnEinlesen_Click(object sender, EventArgs e)
        {
            if (Page.IsValid)
            {

                lblInfo.Text = "Es wird etwa 1,5 min dauern";
                SaveCurrentFormat();
                ExcelFileReader();
                


            }
        }

        private void SaveCurrentFormat()
        {
            string sqlcmd = $"UPDATE eg_lastformat SET Anzahl_Konsumenten = '{txtAnzCon.Text}', Anzahl_Erzeuger = '{txtAnzGen.Text}', Anzahl_Spalten_K = '{txtConCol.Text}', Anzahl_Spalten_E = '{txtGenCol.Text}', Tabellennummer = '{txtNumTab.Text}'";
            DataBase12 dataBase = new DataBase12(connStrg);
            dataBase.ExecuteNonQuery(sqlcmd);
        }

        public bool ExcelFileReader()
        {
            DataBase12 dataBase = new DataBase12(connStrg);
            dataBase.Open();

            XLWorkbook workbook = new XLWorkbook(fuReport.FileContent);
            int numTab = Convert.ToInt32(txtNumTab.Text);
            var ws1 = workbook.Worksheet(numTab);

            int anzCon = Convert.ToInt32(txtAnzCon.Text);
            int anzGen = Convert.ToInt32(txtAnzGen.Text);
            int anzColCon = Convert.ToInt32(txtConCol.Text);
            int anzColGen = Convert.ToInt32(txtGenCol.Text);

            //Wenn die Tabelle angezeigt werden soll muss zuerst die erste Reihe durchnummeriert werden.
            //int colums = Convert.ToInt32(txtAnzCon.Text) * Convert.ToInt32(txtConCol.Text) + Convert.ToInt32(txtAnzGen.Text) * Convert.ToInt32(txtGenCol.Text) + 4;
            //for (int i = 1; i < colums; i++)
            //{
            //    ws1.Cell(1, i).Value = i;
            //}
            //ShowData();


            // ein datensatz = Wer , Wann, wieviel

            //Zeitraum
            DateTime von = Convert.ToDateTime(ws1.Cell(5, 2).Value);
            DateTime bis = Convert.ToDateTime(ws1.Cell(6, 2).Value);
            string vonn = von.ToString("yyyy-MM-dd HH:mm:ss");
            string biss = bis.ToString("yyyy-MM-dd HH:mm:ss");

            string sqlCmd = $"SELECT Count(DatumVon) FROM eg_monatlicherverbrauch Where DatumVon = '{vonn}'";
            int doubleEntrie = Convert.ToInt32(dataBase.RunQueryScalar(sqlCmd));


            //if(von.Month != bis.Month) 




            if (doubleEntrie > 0)
            {
                lblInfo.Text = $"Dieser Monat mit der Datei '{fuReport.FileName}' wurde bereits in der Datenbank gespeichert";
                return false;
            }
            else
            {
                //mit Insert befehl das datum reinschreiben  -- Gemacht

                // sqlCmd = $"INSERT INTO eg_monatlicherverbrauch (DatumVon, DatumBis) VALUES ('{vonn}','{biss}')";
                //  dataBase.ExecuteNonQuery(sqlCmd);

                //sqlCmd = $"INSERT INTO eg_monatlicheerzeugung (DatumVon, DatumBis) VALUES ('{vonn}','{biss}')";
                //dataBase.ExecuteNonQuery(sqlCmd);


                //Verbraucher
                for (int i = 0; i < anzCon; i++)
                {
                    string meteringPointID = ws1.Cell(2, anzColCon + 1 + i * anzColCon).Value.ToString(); //MeteringpointID

                    sqlCmd = $"Select MitgliedsID From eg_mitglieder Where MeteringpointID = '{meteringPointID}'"; // MitgliedID

                    DataTable dt = dataBase.RunQuery(sqlCmd);
                    string mitgliedID = dt.Rows[0][0].ToString();

                    string total = ws1.Cell(11, anzColCon + 1 + i * anzColCon).Value.ToString();    // Es müssen noch doppelte einträge überprüft werden
                    total = total.Replace(',', '.');

                    //MIt update und where die daten einfügen
                    sqlCmd = $"INSERT INTO eg_monatlicherverbrauch (Mitglied, DatumVon, DatumBis, Verbrauch_in_kWh) VALUES ('{mitgliedID}','{vonn}','{biss}','{total}')";
                    dataBase.ExecuteNonQuery(sqlCmd);
                }

                //Generator
                for (int i = 0; i < anzGen; i++)
                {
                    string meteringPointID = ws1.Cell(2, anzColCon * anzCon + 1 + i * anzColGen + anzColGen).Value.ToString();

                    sqlCmd = $"Select MitgliedsID From eg_mitglieder Where MeteringpointID = '{meteringPointID}'"; // MitgliedID

                    DataTable dt = dataBase.RunQuery(sqlCmd);
                    string mitgliedID = dt.Rows[0][0].ToString();

                    string total = ws1.Cell(11, anzColCon * anzCon + 1 + i * anzColGen + anzColGen).Value.ToString();
                    total = total.Replace(',', '.');

                    //MIt update und where die daten einfügen
                    sqlCmd = $"INSERT INTO eg_monatlicheerzeugung (Mitglied, DatumVon, DatumBis, Erzeugung_in_kWh) VALUES ('{mitgliedID}','{vonn}','{biss}','{total}')";
                    dataBase.ExecuteNonQuery(sqlCmd);
                }



                //verbrauchsdaten und euzeugerdaten
                /*
                int j = 12;
                while (ws1.Cell(j, 1).Value.ToString() != "")
                {
                    DateTime quaterHour = Convert.ToDateTime(ws1.Cell(j, 1).Value);
                    string quaterHourString = quaterHour.ToString("yyyy-MM-dd HH:mm:ss");

                    sqlCmd = $"INSERT INTO eg_verbrauchsdaten VALUES ('{quaterHourString}'";
                    for (int i = 0; i < anzCon; i++)
                    {
                        string currentDataValue = ws1.Cell(j, anzColCon + 1 + i * anzColCon).Value.ToString();
                        currentDataValue = currentDataValue.Replace(',', '.');

                        sqlCmd += $" ,'{currentDataValue}'";
                    }
                    sqlCmd += $")";
                    dataBase.ExecuteNonQuery(sqlCmd);


                    sqlCmd = $"INSERT INTO eg_erzeugungsdaten VALUES ('{quaterHourString}'";
                    for (int i = 0; i < anzGen; i++)
                    {
                        string currentDataValue = ws1.Cell(j, anzColCon * anzCon + 1 + i * anzColGen + anzColGen).Value.ToString();
                        currentDataValue = currentDataValue.Replace(',', '.');


                        sqlCmd += $" ,'{currentDataValue}'";
                    }
                    sqlCmd += $");";
                    dataBase.ExecuteNonQuery(sqlCmd);

                    j++;
                }                                                  
            
                */
            }
            dataBase.Close();
            lblInfo.Text = "Datei eingelesen";
            return true;
        }

        private void ShowData()
        {
            ////Daten Anzeigen
            //DataTable dt = new DataTable();
            //bool firstRow = true;

            //foreach (IXLRow row in ws1.Rows())
            //{
            //    while (firstRow)

            //    {
            //        foreach (IXLCell cell in row.Cells())
            //        {
            //            dt.Columns.Add(cell.Value.ToString());
            //        }

            //        dt.Columns.Add(new DataColumn("OK", typeof(string)));
            //        //dt.Columns.Add(new DataColumn("Datum", typeof(string)));
            //        firstRow = false;
            //    }

            //    dt.Rows.Add();
            //    int i = 0;
            //    foreach (IXLCell cell in row.Cells())
            //    {
            //        dt.Rows[dt.Rows.Count -1 ][i] = cell.Value.ToString();
            //        i++;
            //    }

            //}
            //dgvData.DataSource = dt;
            //dgvData.DataBind();
        }

        protected void btnLastFormat_Click(object sender, EventArgs e)
        {
            string sqlcmd = $"Select * From eg_lastformat";
            DataBase12 dataBase = new DataBase12(connStrg);
            DataTable dt = dataBase.RunQuery(sqlcmd);
            txtAnzCon.Text = dt.Rows[0][0].ToString();
            txtAnzGen.Text = dt.Rows[0][1].ToString();
            txtConCol.Text = dt.Rows[0][2].ToString();
            txtGenCol.Text = dt.Rows[0][3].ToString();
            txtNumTab.Text = dt.Rows[0][4].ToString();
        }

        protected void cv_CheckFields_ServerValidate(object source, ServerValidateEventArgs args)
        {
            
            try
            {

                int numTab = Convert.ToInt32(txtNumTab.Text);
                int anzCon = Convert.ToInt32(txtAnzCon.Text);
                int anzGen = Convert.ToInt32(txtAnzGen.Text);
                int anzColCon = Convert.ToInt32(txtConCol.Text);
                int anzColGen = Convert.ToInt32(txtGenCol.Text);
            }
            catch (Exception ex)
            { args.IsValid = false; }
        }

        public void ReadData()
            {
                


            }
    }
}