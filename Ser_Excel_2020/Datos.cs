using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;//Ficheros
using System.Collections;//ArrayList

namespace Ser_Excel_2020
{
    class Datos
    {
        #region Variables Privadas
        private string RutaAplicacion = System.AppDomain.CurrentDomain.BaseDirectory.ToString();
        private LOG log;
        private string narchivo;
        #endregion

        #region Variables Publicas
        public ArrayList RUTAS = new ArrayList();
        public string CargaCarpetaRaiz;
        public string NumSession;
        public ArrayList VecContenido = new ArrayList();
        public int InicioTrama = 0;
        #endregion
        //public Datos(string nombreArchivo)
        public Datos()
        {
            string rtaLog = RutaAplicacion.Replace(@"Ser_excelNV\SIIFNET\", "Documentos\\LOGS\\");
            log = new LOG(rtaLog, "Ser_Excel");
            //narchivo = nombreArchivo;
            CargarRutas(RutaAplicacion);
            CargaCarpetaRaiz = carpetaRaiz();
            //ConsumirArchivo(narchivo);
            //IniTrama();

        }

        private void CargarRutas(string RutaAplicacion)
        {
            try
            {
                string RutaRutas = Path.Combine(RutaAplicacion, "Rutas.txt");
                if (File.Exists(RutaRutas))
                {
                    StreamReader LecturaRutas = new StreamReader(RutaRutas);
                    while (!LecturaRutas.EndOfStream)
                    {
                        RUTAS.Add(LecturaRutas.ReadLine().ToString());
                    }
                    if (RUTAS.Count <= 0)
                    {
                        log.EscribeLog("El archivo de rutas esta vacio" + "___Rutas:" + RutaRutas);
                    }
                }
                else
                {
                    log.EscribeLog("No se encontro el archivo de rutas" + "___Ruta:" + RutaRutas);
                }
            }
            catch (Exception ex)
            {
                log.EscribeLog("Ocurrio un error mientras se accedia al archivo de rutas", ex.Message.ToString());
            }
        }
        //Se obtiene la carpeta Raíz del ambiente
        private string carpetaRaiz()
        {
            string carpetaR = "";
            int posicion;
            posicion = RutaAplicacion.LastIndexOf("\\");
            carpetaR = RutaAplicacion.Substring(0, posicion);
            posicion = carpetaR.LastIndexOf("\\");
            carpetaR = carpetaR.Substring(0, posicion);
            posicion = carpetaR.LastIndexOf("\\");
            carpetaR = carpetaR.Substring(0, posicion) + "\\";

            return carpetaR;
        }
    }
}
