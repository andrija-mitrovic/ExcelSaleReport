using Advantage.Data.Provider;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelSaleReport.Data
{
    public class DbConnection
    {
        private static string _pathDbf = Path.GetFullPath(Path.Combine(Application.StartupPath, "..\\..\\..\\"));
        private string _connectionString = "Data Source=" + _pathDbf + ";servertype=local;TableType=NTX;";
        private AdsConnection _connection;
        private AdsDataAdapter _dataAdapter;
        private AdsCommand _command;
          
        public DbConnection()
        {
            _dataAdapter = new AdsDataAdapter();
            _connection = new AdsConnection(_connectionString);
        }

        private AdsConnection OpenConnection()
        {
            if (_connection.State == ConnectionState.Closed || _connection.State == ConnectionState.Broken)
            {
                _connection.Open();
            }
            return _connection;
        }

        public DataTable ExecuteSelectQuery(string query)
        {
            _command = new AdsCommand();
            DataTable dt=null;
            DataSet ds = new DataSet();

            try
            {
                using (_command = new AdsCommand(query, OpenConnection()))
                {
                    _command.CommandTimeout = 0;
                    _dataAdapter = new AdsDataAdapter(_command);
                    ds = new DataSet();
                    _dataAdapter.Fill(ds);
                    dt = ds.Tables[0];
                }
            }
            catch(AdsException ex)
            {
                MessageBox.Show("Error - ExecuteSelectQuery - Query: "+ query +"\n Exception: " + ex.Message);
            }

            return dt;
        }

        public bool ExecuteInsertQuery(string query)
        {
            _command = new AdsCommand();

            try
            {
                using(_command = new AdsCommand(query, OpenConnection()))
                {
                    _command.ExecuteNonQuery();
                }
            }
            catch(AdsException ex)
            {
                MessageBox.Show("Error - ExecuteInsertQuery - Query: " + query + "\n Exception: " + ex.Message);
                return false;
            }

            return true;
        }

        public bool ExecuteUpdateQuery(string query)
        {
            _command = new AdsCommand();

            try
            {
                using (_command = new AdsCommand(query, OpenConnection()))
                {
                    _command.ExecuteNonQuery();
                }
            }
            catch(AdsException ex)
            {
                MessageBox.Show("Error - ExecuteUpdateQuery - Query: " + query + "\n Exception: " + ex.Message);
                return false;
            }

            return true;
        }

        public bool ExecuteDeleteQuery(string query)
        {
            _command = new AdsCommand();

            try
            {
                using (_command = new AdsCommand(query, OpenConnection()))
                {
                    _command.ExecuteNonQuery();
                }
            }
            catch(AdsException ex)
            {
                MessageBox.Show("Error - ExecuteDeleteQuery - Query: " + query + "\n Exception: " + ex.Message);
                return false;
            }

            return true;
        }
    }
}
