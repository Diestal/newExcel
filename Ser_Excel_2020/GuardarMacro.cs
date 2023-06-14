using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using System.Windows;
using OfficeOpenXml;//Librería EPPlus
using OfficeOpenXml.Table;
using OfficeOpenXml.Drawing;
using System.Drawing;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;//Librería de Excel

namespace Ser_Excel_2020
{
    class GuardarMacro
    {
        #region VARIABLES LIBRERÍA EPPLUS
        private ExcelPackage AppExcel;//Librería EPPlus - Instancia para crear el Excel
        private ExcelWorkbook libroXls;
        #endregion

        #region VARIABLES CLASES
        private LOG log;
        private Datos CargaDatos;
        #endregion

        #region VARIABLES FICHEROS
        private FileInfo existeArchivo;
        #endregion

        #region VARIABLES GENERALES
        private string RutaAplicacion = System.AppDomain.CurrentDomain.BaseDirectory.ToString();
        private string trama = "", plantilla, nombreXLS;
        private int linea = 1;
        private int numero;
        private StreamReader f1;
        #endregion

        public GuardarMacro()
        {
            string rtaLog = RutaAplicacion.Replace(@"SIIFNET\", "Documentos\\LOGS\\");
            CargaDatos = new Datos();
            log = new LOG(rtaLog, "Ser_Excel");
        }
        public void guardarReporte(string ruta_Timelog)
        {
            Random rnd = new Random();
            numero = rnd.Next(1000, 9999);

            /*
             //////////////////////////////////////////////////////////////////////
             Archivo MODIFICAEX, para pruebas
            ------------------------------------------------------------------
             Plantilla          ->  CRINCREQ
             NombreExcel        ->  M95375.XLS
             Trama para merchar ->  PRUEBA 12345|PRUEBA 6789|||||7|1|9|
             -----------------------------------------------------------------
             ///////////////////////////////////////////////////////////////////////
             */

             /* DOCUMENTOS\ARCHIVOS\MODIFICAEX.TXT */
            if (File.Exists(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "modificaex.txt"))
            {
                try
                {
                    if (File.Exists(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "modificaex" + numero + ".txt"))
                        File.Delete(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "modificaex" + numero + ".txt");

                    File.Move(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "modificaex.txt", CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "modificaex" + numero + ".txt");
                }
                catch (Exception ex)
                {
                    log.EscribeLog("Error al renombrar el archivo : " + CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "modificaex.txt como: [" + CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "modificaex" + numero + ".txt]", ex.ToString(), true);
                }

                try
                {
                    f1 = new StreamReader(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "modificaex" + numero + ".txt");
                    //RECORRE ARCHIVO MODIFICAEX
                    while (!f1.EndOfStream)
                    {
                        switch (linea)
                        {
                            case 1:
                                plantilla = f1.ReadLine().Trim();
                                break;
                            case 2:
                                nombreXLS = f1.ReadLine().Trim();
                                break;
                            default:
                                trama = trama + f1.ReadLine().Trim();
                                break;
                        }
                        linea += 1;
                    }
                    f1.Close();
                }
                catch (Exception ex)
                {
                    log.EscribeLog("Error al abrir y cerrar el archivo: [ " + CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "modificaex" + numero + ".txt ]", ex.ToString(), true);
                }

                //ELIMINA ARCHIVO MODIFICAEX1234.TXT
                try
                {
                    File.Delete(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "modificaex" + numero + ".txt");
                }
                catch (Exception ex)
                {
                    log.EscribeLog("Error al eliminar: [ " + CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "modificaex" + numero + ".txt ]", ex.ToString(), true);
                }

                guardar(plantilla, nombreXLS, trama, ruta_Timelog);

            }

        }

        //METODO PARA MERCHAR LA INFORMACIÓN DEL MACRO
        public void guardar(string plantilla, string nombreXLS, string trama, string ruta_Timelog)
        {
            DateTime Hora;
            string archivoXLS = "";
            string[] vec;

            archivoXLS = CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + nombreXLS;
            vec = trama.Split('|');

            if(File.Exists(archivoXLS))
            {
                try
                {
                    existeArchivo = Utilidades.ObtieneInfoArchivo(archivoXLS, false);//false -> si no se desea eliminar el archivo
                    using (AppExcel = new ExcelPackage(existeArchivo))
                    {
                        try
                        {

                            string Varreemplazar = "";
                            AppExcel.Workbook.Worksheets.First();//SE SELECCIONA LA PRIMERA HOJA
                            for (int i = 0; i < vec.Length - 1; i++)
                            {
                                libroXls = AppExcel.Workbook;//LOS CUADROS O CAJAS DE NOMBRES RECAEN ES EN EL LIBRO, MAS NO EN LAS HOJAS
                                Varreemplazar = "_ENV" + (i+1).ToString("000");
                                libroXls.Names[Varreemplazar].Value = vec[i];
                            }
                        }
                        catch (Exception ex)
                        {
                            log.EscribeLog("Error al reemplazar variables ", ex.ToString(), true);
                        }

                        AppExcel.Compression = CompressionLevel.BestSpeed;

                        //SE GUARDA EL ARCHIVO FINAL
                        var ArchivoExcel = Utilidades.ObtieneInfoArchivo(archivoXLS, false);
                        AppExcel.SaveAs(ArchivoExcel);

                        //Controlador de tiempo en generar el reporte
                        Hora = DateTime.Now;
                        StreamWriter timelog = new StreamWriter(ruta_Timelog, true, Encoding.UTF8);
                        timelog.WriteLine("***************************************************************************************");
                        timelog.WriteLine("Hora de Finalización:    " + Hora.ToString("F"));//F - Friday, February 27, 2009 12:12:22 PM;
                        timelog.Flush();
                        timelog.Close();

                        System.Environment.Exit(0);
                    }
                }
                catch(Exception ex)
                {
                    log.EscribeLog("Error al crear Objeto de Excel ", ex.ToString(), true);
                }
            }
            else
            {
                log.EscribeLog("El Archivo : [ " + CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[3].ToString() + plantilla + " ] no existe", "", true);
            }

        }
    }
}
