using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Ser_Excel_2020
{
    class Utilidades
    {
        static DirectoryInfo _CrearDirectorio = null;
        public static DirectoryInfo CrearDirectorio
        {
            get
            {
                return _CrearDirectorio;
            }
            set
            {
                _CrearDirectorio = value;
                if (!_CrearDirectorio.Exists)
                {
                    _CrearDirectorio.Create();
                }
            }
        }
        public static FileInfo ObtieneInfoArchivo(string file, bool deleteIfExists = true)
        {
            //var fi = new FileInfo(CrearDirectorio.FullName + Path.DirectorySeparatorChar + file);
            var fi = new FileInfo(file);
            if (deleteIfExists && fi.Exists)
            {
                fi.Delete();  // Asegura que se haya creado un nuevo libro Excel
            }
            return fi;
        }
        internal static DirectoryInfo ObtieneInfoDirectorio(string directory)
        {
            var di = new DirectoryInfo(_CrearDirectorio.FullName + Path.DirectorySeparatorChar + directory);
            if (!di.Exists)
            {
                di.Create();
            }
            return di;
        }
    }
}
