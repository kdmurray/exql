using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace ExQL
{
    class ExcelManager
    {
        static string[] _validExtensions = new string[] { "xls", "xlsx" };

        OleDbConnection _conn;
        OleDbCommand _cmd;
        OleDbDataAdapter _da;

        public ExcelManager(string SourceFile)
        {
            _conn = new OleDbConnection(GenerateConnectionString(SourceFile));
            _cmd = new OleDbCommand();
            _cmd.Connection = _conn;
            _da = new OleDbDataAdapter();
        }

        public DataTable Execute(string Query)
        {
            _cmd.CommandText = Query;
            _da.SelectCommand = _cmd;
            
            DataSet ds = new DataSet();
            _da.Fill(ds);
            
            return ds.Tables[0];
        }

        string GenerateConnectionString(string SourceFile)
        {
            StringBuilder conString = new StringBuilder();
            conString.Append("Provider=Microsoft.ACE.OLEDB.12.0;");
            conString.Append("Data Source=");
            conString.Append(SourceFile);
            conString.Append(";");
            conString.Append("Extended Properties=\"Excel 12.0 Xml;HDR=Yes;\"");
            return conString.ToString();
        }

        public static ExcelFileType GetFileType(string SourceFile)
        {
            string extension = SourceFile.Substring(SourceFile.LastIndexOf(".") + 1).ToLower();

            if (_validExtensions.Contains(extension))
            {
                return (ExcelFileType)Enum.Parse(typeof(ExcelFileType), extension.ToUpper());
            }

            return ExcelFileType.INVALID;
        }
    }
}
