
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleTeste
{
    public class Program
    {
        static void Main(string[] args)
        {
            char[] letras = new char[9];
            string palavra = "ola mundo";

            for (int i = 0; i < palavra.Length; i++)
            {
                letras[i] = palavra[i];
            }

            for (int i = 0; i < letras.Length; i++)
            {
                // Console.Write(letras[i]);
            }
            Console.ReadKey();
        }

        static void OlaMundo()
        { }

        static int SomaNumeros(int numA, int numB)
        {
            int soma = numA + numB;
            Console.WriteLine(soma);

            return soma;
        }
        public static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }
    }
}
