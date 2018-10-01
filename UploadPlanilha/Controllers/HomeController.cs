using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using System.IO;

namespace UploadPlanilha.Controllers
{
    public class HomeController : Controller
    {
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["BANCO"].ConnectionString);

        OleDbConnection Econ;

        private void ExcelConn(string filepath)

        {
            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", filepath);

            Econ = new OleDbConnection(constr);
        }

        private void InsertExceldata(string fileepath, string filename)

        {

            string fullpath = Server.MapPath("/excelfolder/") + filename;

            ExcelConn(fullpath);

            string query = string.Format("Select * from [{0}]", "Coluna$");

            OleDbCommand Ecom = new OleDbCommand(query, Econ);

            Econ.Open();

            DataSet ds = new DataSet();

            OleDbDataAdapter oda = new OleDbDataAdapter(query, Econ);

            Econ.Close();

            oda.Fill(ds);

            DataTable dt = ds.Tables[0];

            SqlBulkCopy objbulk = new SqlBulkCopy(con);

            if (dt.Rows.Count > 0)
            {
                DataColumn newColumn = new DataColumn("DataAlteracao", typeof(System.DateTime));
                newColumn.DefaultValue = DateTime.Now;
                dt.Columns.Add(newColumn);
            }


            objbulk.DestinationTableName = "RegistroMovimentacaoEquipamento";
            objbulk.ColumnMappings.Add("NrPatrimonio","NrPatrimonio");
            objbulk.ColumnMappings.Add("NrSerie","NrSerie");
            objbulk.ColumnMappings.Add("TipoId","TipoId");
            objbulk.ColumnMappings.Add("MarcaId","MarcaId");
            objbulk.ColumnMappings.Add("Descricao","Descricao");
            objbulk.ColumnMappings.Add("ModeloId","ModeloId");
            objbulk.ColumnMappings.Add("StatusId","StatusId");
            objbulk.ColumnMappings.Add("FilialId","FilialId");
            objbulk.ColumnMappings.Add("DepartamentoId","DepartamentoId");
            objbulk.ColumnMappings.Add("ColaboradorId","ColaboradorId");
            objbulk.ColumnMappings.Add("Justificativa","Justificativa");
            objbulk.ColumnMappings.Add("ColaboradorIdAlteracao","ColaboradorIdAlteracao");
            objbulk.ColumnMappings.Add("ChamadoId","ChamadoId");
            objbulk.ColumnMappings.Add("NrNota","NrNota");
            objbulk.ColumnMappings.Add("NrNotaSaida","NrNotaSaida");
            objbulk.ColumnMappings.Add("DataAlteracao","DataAlteracao");

            con.Open();

            objbulk.WriteToServer(dt);

            con.Close();

        }


        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            string filename = DateTime.Now.Day+"-"+DateTime.Now.Month+"-"+DateTime.Now.Year 
                + "-"+DateTime.Now.Hour + "-" + DateTime.Now.Minute + "-" + DateTime.Now.Second + Path.GetExtension(file.FileName);

            string filepath = "/excelfolder/" + filename;

            file.SaveAs(Path.Combine(Server.MapPath("/excelfolder"), filename));

            InsertExceldata(filepath, filename);
            return View();
        }


    }
}