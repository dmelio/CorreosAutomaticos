using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace CorreosAutomaticos
{
    class Data
    {
        Conexion conexion;
        public Data()
        {
            conexion = new Conexion();
        }

        public DataSet MostrarData()
        {
            SqlCommand sentencia = new SqlCommand("exec dbo.Proc_Ctrl_Puntos");
            return conexion.EjecutarSentencia(sentencia);
        }

    }
}
