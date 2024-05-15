using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Cryptography.X509Certificates;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Security;
using System.Net;
using System.Reflection.Emit;
using Microsoft.Office.Interop.Excel;
using System.Data.Odbc;



namespace CorreosAutomaticos
{
    public partial class Form1 : Form
    {

        private string correoEmisor = "d.melio.escobar@gmail.com";
        private string passEmisor = "udtlcqwqtgxgwwbc";
        private string myAlias = "Daniel Melio";
        private string[] myAdjuntos;
        private string[] destinatarios;
        private MailMessage mCorreo;
        string rutaArchivo = "C:\\Users\\56961\\Desktop\\Adasme\\ReporteSW\\Prueba_informe.xlsx";

        Data datos;
        private Timer timer;
        Data oData;
        public Form1()
        {
            oData = new Data();
            InitializeComponent();
            dataGridView1.DataSource = oData.MostrarData().Tables[0];
            timer = new Timer();
            timer.Interval = 4 * 60 * 60 * 1000; //4 horas 
            timer.Tick += timer1_Tick;
            timer.Start();
        }


        private void CrearCuerpoCorreo()
        {
            mCorreo = new MailMessage();
            mCorreo.From = new MailAddress(correoEmisor, myAlias, System.Text.Encoding.UTF8);
            //mCorreo.To.Add("daniel.melio@outlook.com");
            List<string> destinatarios = new List<string>
             {
               "daniel.melio@outlook.com"
             };
            foreach (var destinatario in destinatarios)
            {
                mCorreo.To.Add(destinatario);
            };

            //for (int i = 0; i < destinatarios.Length; i++)
            //{
            //    mCorreo.To.Add(destinatarios[i]);
            //}
            DateTime marcaTiempo = DateTime.Now;
            string NombreArchivoAct = $"Prueba_informe{marcaTiempo.ToString("yyyyMMdd_HHmmss")}.xlsx";
            ActualizarPlanilla(rutaArchivo,NombreArchivoAct);
            mCorreo.Attachments.Add(new Attachment($"C:\\Users\\56961\\Desktop\\Adasme\\ReporteSW\\{NombreArchivoAct}"));
            mCorreo.Subject = "Reporte";
            //mCorreo.Body = richTextBox1.Text.Trim();
            mCorreo.IsBodyHtml = true;
            mCorreo.Priority = MailPriority.High;

            //adjuntos
            //for (int i = 0; i < myAdjuntos.Length; i++)
            //{
            //    mCorreo.Attachments.Add(new Attachment(myAdjuntos[i]));
            //}

        }
        
        private void ActualizarTablaDinamica(Excel.Workbook workbook)
        {
            Excel.PivotTable pivotTable = (Excel.PivotTable)workbook.Sheets["Detalle_y_grafico"].PivotTables("TablaDinámica1");
            pivotTable.RefreshTable();
        }
        private void ActualizarPlanilla(string rutaArchivo,string NombreArchivoAct)
        {
            datos = new Data();
            
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook wb = excelApp.Workbooks.Open(rutaArchivo, 0, false);
            Excel.Worksheet worksheet = wb.Sheets[1];
            worksheet.Cells[1, 1] = "Fecha_Control";
            worksheet.Cells[1, 2] = "Hora_Control";
            worksheet.Cells[1, 3] = "Punto";
            worksheet.Cells[1, 4] = "Frame";
            worksheet.Cells[1, 5] = "Diff_Seg1";
            worksheet.Cells[1, 6] = "Diff_Seg2";
            worksheet.Cells[1, 7] = "Diff_Seg3";
            worksheet.Cells[1, 8] = "Diff_Seg4";
            worksheet.Cells[1, 9] = "Diff_Seg5";
            worksheet.Cells[1, 10] = "Diff_Seg6";
            worksheet.Cells[1, 11] = "Diff_Seg7";
            worksheet.Cells[1, 12] = "Diff_Seg8";
            worksheet.Cells[1, 13] = "Diff_Seg9";
            int maxRow = 20;
            int maxCol = 13;

            for (int i = 1; i <= maxRow; i++)
            {
                for (int j = 1; j <= maxCol; j++)
                {
                    worksheet.Cells[i + 5, j + 1].Value = null;
                }
            }

            for (int i = 0; i < datos.MostrarData().Tables[0].Rows.Count; i++)
            {
                for (int j = 0; j < datos.MostrarData().Tables[0].Columns.Count; j++)
                {
                    worksheet.Cells[i + 5, j + 1].Value = datos.MostrarData().Tables[0].Rows[i][j];
                }
            }
            ActualizarTablaDinamica(wb);
            
            wb.SaveAs($"C:\\Users\\56961\\Desktop\\Adasme\\ReporteSW\\{NombreArchivoAct}");
            wb.Close(false);
            //excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
        //private void EjecutarPlanilla(string rutaArchivo)
        //{
        //    Excel.Application excelApp = new Excel.Application();
        //    excelApp.Visible = true; 
        //    Excel.Workbook workbook = excelApp.Workbooks.Open(rutaArchivo);
            //Excel.Worksheet worksheet = workbook.Worksheets["Detalle_y_grafico"];
            //worksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
        //    workbook.Save();
        //    workbook.Close();

        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        //}

        private void Enviar()
        {
            try
            {
                SmtpClient smtp = new SmtpClient();
                smtp.UseDefaultCredentials = false;
                smtp.Port = 25;
                smtp.Host = "smtp.gmail.com";
                smtp.Credentials = new System.Net.NetworkCredential(correoEmisor, passEmisor);
                ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
                {
                    return true;
                };
                smtp.EnableSsl = true;
                smtp.Send(mCorreo);
                MessageBox.Show("Enviado");
                //TxtLblAdj.Text = "...";
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void AdjuntarArchivos()
        {
            try
            {
                var names = "";
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Multiselect = true;
                ofd.Title = "Adjuntar Archivos";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    myAdjuntos = ofd.FileNames;
                }
                for (int i = 0; i < myAdjuntos.Length; i++)
                {
                    names += myAdjuntos[i] + "\n";
                }
                TxtLblAdj.Text = names;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            
        }
        private void BtnAdjuntar_Click(object sender, EventArgs e)
        {
            AdjuntarArchivos();
        }

        private void BtnEnviar_Click(object sender, EventArgs e)
        {
            CrearCuerpoCorreo();
            Enviar();
        }

        
        private void BtnActualizar_Click(object sender, EventArgs e)
        {
            //EjecutarPlanilla("C:\\Users\\56961\\Documents\\prueba.xlsx");
            Conexion conexion = new Conexion();
            //ActualizarPlanilla(rutaArchivo);
            //MessageBox.Show("Conectado .. " + conexion.PruebaConectar());
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            dataGridView1.DataSource = oData.MostrarData().Tables[0];
            //Conexion conexion = new Conexion();
            //ActualizarPlanilla(rutaArchivo);
            CrearCuerpoCorreo();
            //MessageBox.Show("Conectado .. " + conexion.PruebaConectar());
            Enviar();
        }
    }
}
