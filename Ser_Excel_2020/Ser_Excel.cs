using System;
using System.Text;
using System.IO;

namespace Ser_Excel_2020
{
    public class Ser_Excel
    {
        #region Variables
        private string RutaAplicacion = System.AppDomain.CurrentDomain.BaseDirectory.ToString();
        private string[] _param;
        private LOG log;
        private Datos CargaDatos;
        private GuardarMacro guardaExcelMacro;
        private CrearExcel reporteExcel;
        public string rutaPDF;
        object misValue = System.Reflection.Missing.Value;
        #endregion
        public Ser_Excel(string[] param)
        {
            StreamWriter timelog;
            /* 
            d    2/27/2009
            D    Friday, February 27, 2009
            f    Friday, February 27, 2009 12:11 PM
            F    Friday, February 27, 2009 12:12:22 PM
            g    2/27/2009 12:12 PM
            G    2/27/2009 12:12:22 PM
            m    February 27
            M    February 27
            o    2009-02-27T12:12:22.1020000-08:00
            O    2009-02-27T12:12:22.1020000-08:00
            s    2009-02-27T12:12:22
            t    12:12 PM
            T    12:12:22 PM
            u    2009-02-27 12:12:22Z
            U    Friday, February 27, 2009 8:12:22 PM
            y    February, 2009
            Y    February, 2009
        */
            DateTime Hora;
            string ruta_Timelog = "";
            string rtaLog = RutaAplicacion.Replace(@"Ser_excelNV\SIIFNET\", "Documentos\\LOGS\\");
            ruta_Timelog = RutaAplicacion + "TimeLog.txt";

            param = Environment.GetCommandLineArgs();

            //------------------------------------
            //--Para Pruebas, Excel normal
            //------------------------------------
            //param = new string[]{ @"C:\Users\robert.monterrosa\source\repos\Ser_Excel_2020\Ser_Excel_2020\bin\Debug\Ser_Excel.exe", "P00000000057649" };//El separador no se visualiza en este editor
            /*param = new string[] { @"C:\Users\brayan.milian\Documents\Informacion\SIIFGIT\Interactuar\ConsoleApp-FormApp\Ser_Excel_2020\Ser_Excel_2020\bin\Debug\Ser_Excel.exe", "P00001860118602" };*///El separador no se visualiza en este editor
            //_param = param;
            //------------------------------------
            //--Para Pruebas, Excel con Macros
            //------------------------------------
            //param = new string[]{ "Ruta Ser_Excel", "X"};//Para pruebas
            CargaDatos = new Datos();
            
            //***************************************************************
            log = new LOG(rtaLog, "Ser_Excel");

            //Controlador de tiempo en generar el reporte
            Hora = DateTime.Now;
            timelog = new StreamWriter(ruta_Timelog, true, Encoding.UTF8);
            timelog.WriteLine("***************************************************************************************");
            timelog.WriteLine("Hora de Inicio:    " + Hora.ToString("F"));//F - Friday, February 27, 2009 12:12:22 PM;
            timelog.Flush();
            timelog.Close();


            if(param.Length > 0)
            {
                if (param[1].Substring(0, 1) == "X")
                {
                    //EXCEL CON MACROS
                    guardaExcelMacro = new GuardarMacro();
                    guardaExcelMacro.guardarReporte(ruta_Timelog);
                }
                else
                {
                    //CREACIÓN DE REPORTES MAS UTILIZADO
                    reporteExcel = new CrearExcel();
                    rutaPDF = reporteExcel.crear_reporte(param, ruta_Timelog).Result;
                }
            }
            else
            {
                log.EscribeLog("El Parametro Nº de Session esta Vacio", "", true);
            }

        }
    }
}
