﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ser_Excel_2020
{
    class Program
    {
        //(Single Thread Apartment, STA)
        [STAThread]
        static void Main(string[] args)
        {
            var parametro = new string[]{ @"", "" };//El separador no se visualiza en este editor

            Ser_Excel ser_Excel = new Ser_Excel(parametro);
        }
    }
}
