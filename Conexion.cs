using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorreosAutomaticos
{
    class Conexion
    {
        public string CadenadeConexion = "Data Source= 138.121.170.38,1434; Initial Catalog=Automotriz; User ID=DevUser; Password=1!querty";
        SqlConnection Connect;

        public SqlConnection EstablecenConexion()
        {
            this.Connect = new SqlConnection(this.CadenadeConexion);
            return this.Connect;
        }
        public bool PruebaConectar()
        {
            try
            {
                
                SqlCommand Comando = new SqlCommand();

                Comando.CommandText = "Select * from Clientes";
                Comando.Connection = this.EstablecenConexion();
                Connect.Open();
                Comando.ExecuteNonQuery();
                Connect.Close();
                return true;
            }
            catch 
            {
                return false;
            }

        }

        public DataSet EjecutarSentencia(SqlCommand sqlComando)
        {
            DataSet ds = new DataSet();
            SqlDataAdapter Adaptador = new SqlDataAdapter();

            try
            {
                SqlCommand Comando = new SqlCommand();
                Comando = sqlComando;
                Comando.Connection = EstablecenConexion();
                Adaptador.SelectCommand = Comando;
                Connect.Open();
                Adaptador.Fill(ds);
                Connect.Close();
                return ds;

            }
            catch
            {
                return ds;
            }
        }
    }
}
