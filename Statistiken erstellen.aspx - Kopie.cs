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
using Plotly.NET;
using XPlot.Plotly;

using Microsoft.Data.Analysis;

//using Microsoft.AspNetCore.Html;



namespace Excel_Hochladen_und_Einlesen
{
    public partial class Statistiken_erstellen : System.Web.UI.Page
    {
        string connStrg = WebConfigurationManager.ConnectionStrings["AppDbInt"].ConnectionString;
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btn_Statistik_Click(object sender, EventArgs e)
        {
            DataBase12 dataBase = new DataBase12(connStrg);
            dataBase.Open();

            DateTime dateVon = cal_Von.SelectedDate;
            DateTime dateBis = cal_Bis.SelectedDate;
            string verbraucher = txt_Konsument.Text;



            string sqlCmd = $"SELECT Zeit, {verbraucher} FROM eg_verbrauchsdaten";
            DataTable dt = dataBase.RunQuery(sqlCmd);




            
        }
    }
}