using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Ser_Excel_2020
{
    class LOG
    {
        #region Variables de Clase 
        public string RutaLog;
        public string Resultado;
        private bool rutaCorrecta;
        public bool Indica;
        private string NomAplicacion;
        private StreamWriter Escribe;
        #endregion

        #region Constructor
        public LOG(string RtaLog, string NombreAplicacion)
        {
            try
            {
                RtaLog = RtaLog + "\\" + NombreAplicacion + ".txt";
                if (File.Exists(RtaLog))
                {
                    RutaLog = RtaLog;
                    rutaCorrecta = true;
                    Escribe = new StreamWriter(RtaLog, true, Encoding.UTF8);
                    NomAplicacion = NombreAplicacion;
                    Escribe.WriteLine(NombreAplicacion + "_____" + DateTime.Now);
                    Escribe.Close();
                }
                else
                {
                    Escribe = new StreamWriter(RtaLog, true, Encoding.UTF8);
                    Escribe.WriteLine(NombreAplicacion + "_____" + DateTime.Now);
                    rutaCorrecta = true;
                    RutaLog = RtaLog;
                    NomAplicacion = NombreAplicacion;
                    Escribe.Close();
                }
            }
            catch
            {
                rutaCorrecta = false;
                Resultado = "No se pudo comprobar la ruta del log";
                Escribe.Close();
            }
        }
        #endregion

        #region Metodos
        public void EscribeLog(string MensajeLog, string Ex = " ", bool salir = false)
        {
            try
            {
                if (rutaCorrecta)
                {
                    Escribe = new StreamWriter(RutaLog, true);
                    Escribe.WriteLine(DateTime.Now.ToString() + " " + NomAplicacion);
                    Escribe.WriteLine(MensajeLog + " -> Descripcion del Programa: " + Ex);
                    Escribe.WriteLine("********************************************************************************");
                    Escribe.Close();
                }
                else
                {
                    Resultado = "No se escribio el LOG";
                }
            }
            catch
            {
                Resultado = "Error intentando escribir el log, verifique que el archivo no este corrupto";
                Escribe.Close();
            }
            if (salir)
            {
                System.Environment.Exit(0);
            }
        }
        #endregion
    }
}
