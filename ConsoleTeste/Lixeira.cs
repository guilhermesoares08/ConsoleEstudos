using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleTeste
{
    class Lixeira
    {
        public static void FindJairo(IList<Pessoa> listaPessoas)
        {
            if (listaPessoas.Any())
            {
                Pessoa jairoAchado1 = (from item in listaPessoas where item.Name.Contains("Jairo") select item).First();
                if (jairoAchado1 != null)
                {
                    Console.WriteLine($"gol bolinha do {jairoAchado1.Name}");
                }
            }
        }

        public void Menu()
        {


            Console.WriteLine("Enviar arquivos?");
            Console.WriteLine("S - SIM");
            Console.WriteLine("N - NÃO");


            string op = Console.ReadLine();

            switch (op.ToUpper())
            {
                case "S":
                    //readExcel
                    //save to and bcc                    
                    List<string> to = new List<string>();
                    List<string> bcc = new List<string>();
                    //MailHelper mail = new MailHelper(to, bcc);
                    break;
                default:
                    break;
            }
        }


        public static string[] GetRange(string range, Worksheet excelWorksheet)
        {
            Microsoft.Office.Interop.Excel.Range workingRangeCells =
              excelWorksheet.get_Range(range, Type.Missing);


            System.Array array = (System.Array)workingRangeCells.Cells.Value2;
            string[] arrayS = ConvertToStringArray(array);

            return arrayS;
        }

        public static string[] ConvertToStringArray(System.Array values)
        {

            // create a new string array
            string[] theArray = new string[values.Length];

            // loop through the 2-D System.Array and populate the 1-D String Array
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(1, i) == null)
                    theArray[i - 1] = "";
                else
                    theArray[i - 1] = (string)values.GetValue(1, i).ToString();
            }

            return theArray;
        }

        public static void CreateFileExcel(string fileName)
        {
            Excel.Application excel = new Excel.Application();
            if (excel == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }

            var workBook = excel.Workbooks.Add();
            var workSheet = (Excel.Worksheet)excel.ActiveSheet;
            workSheet.Cells[1, "A"] = "Foo";
            workSheet.Cells[1, "B"] = "Bar";

            workBook.SaveAs(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + fileName + DateTime.Now.ToString("ddMMyyyyhhmmss"), Excel.XlFileFormat.xlOpenXMLWorkbook);

        }

        private void CreateFileExcel2(string fileName)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "ID";
            xlWorkSheet.Cells[1, 2] = "Name";
            xlWorkSheet.Cells[2, 1] = "1";
            xlWorkSheet.Cells[2, 2] = "One";
            xlWorkSheet.Cells[3, 1] = "2";
            xlWorkSheet.Cells[3, 2] = "Two";

            xlWorkBook.SaveAs(Directory.GetCurrentDirectory() + "\\" + fileName + DateTime.Now.ToString("ddMMyyyyhhmmss"), Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("Excel file created , you can find the file d:\\csharp-Excel.xls");
        }


        public void Dump()
        {
            //    List<Pessoa> lista = new List<Pessoa>();

            //    for (int i = 0; i < 10; i++)
            //    {
            //        //Pessoa q = new Pessoa();
            //        //q.Nome = "Pessoa" + i;
            //        //q.Cpf = i.ToString();
            //        //q.Idade = i;
            //        //lista.Add(q);
            //    }

            //    foreach (Pessoa pessoa in lista)
            //    {
            //        Console.WriteLine("Nome:" + pessoa.Nome);
            //        Console.WriteLine("Cpf:" + pessoa.Cpf);
            //        Console.WriteLine("Idade:" + pessoa.Idade);
            //        Console.WriteLine("----------");
            //    }

            //    string resposta = "";
            //    int posicao = 0;

            //    Console.WriteLine("Deseja alterar alguem?");
            //    Console.WriteLine("S/N?");
            //    resposta = Console.ReadLine();

            //    if (resposta == "S")
            //    {
            //        Console.WriteLine("qual pessoa? (posicao)");
            //        posicao = Convert.ToInt32(Console.ReadLine());

            //        int propriedade = 0;
            //        Console.WriteLine("Pessoa escolhida: " + lista[posicao].Nome);
            //        Console.WriteLine("qual valor deseja editar?:");
            //        Console.WriteLine("1 - Nome");
            //        Console.WriteLine("2 - Cpf");
            //        Console.WriteLine("3 - Idade");
            //        propriedade = Convert.ToInt32(Console.ReadLine());

            //        if (propriedade == 1)
            //        {
            //            Console.WriteLine("Digite o novo nome");
            //            lista[posicao].Nome = Console.ReadLine();
            //        }
            //        else if (propriedade == 2)
            //        {
            //            Console.WriteLine("Digite o novo cpf");
            //            lista[posicao].Cpf = Console.ReadLine();
            //        }
            //        else if (propriedade == 3)
            //        {
            //            Console.WriteLine("Digite a nova idade:");
            //            lista[posicao].Idade = Convert.ToInt32(Console.ReadLine());
            //        }

            //        Console.WriteLine("Nome:" + lista[posicao].Nome);
            //        Console.WriteLine("Cpf:" + lista[posicao].Cpf);
            //        Console.WriteLine("Idade:" + lista[posicao].Idade);
            //    }


            //string opcao = "";
            //do
            //{
            //    Console.Clear();
            //    string tamanhoDoVetorPeloUsuario = "";
            //    Console.WriteLine("Digite o tamanho que quer o vetor:");
            //    int tamanhoVetorInteiro = 0;
            //    try
            //    {
            //        tamanhoDoVetorPeloUsuario = Console.ReadLine();
            //        tamanhoVetorInteiro = int.Parse(tamanhoDoVetorPeloUsuario);
            //    }
            //    catch
            //    {
            //        Console.ForegroundColor = ConsoleColor.Red;
            //        throw new JaderInputException("o fei tu digitou " + tamanhoDoVetorPeloUsuario + " e fechou o pograma fei");
            //    }

            //    int[] arr = new int[tamanhoVetorInteiro];

            //    for (int i = 0; i < arr.Length; i++)
            //    {
            //        Console.WriteLine("vetor[" + i + "] " + arr[i]);
            //    }

            //    Console.WriteLine("sair? pressione 'S', ou qlquer tecla para continuar");
            //    opcao = Console.ReadLine();

            //} while (opcao.ToLower() != "s");



            //Console.WriteLine("linha:");
            //int rowLength = int.Parse(Console.ReadLine());

            //Console.WriteLine("coluna:");
            //int colLength = int.Parse(Console.ReadLine());

            //int[,] arr = new int[rowLength, colLength];

            //for (int i = 0; i < rowLength; i++)
            //{
            //    for (int j = 0; j < colLength; j++)
            //    {
            //        Console.Write("1");
            //    }
            //    Console.WriteLine();
            //}
        }

        public void JaderExceptionn()
        {
            List<int> listDeRodoes = new List<int>();
            string opcerteza = "";

            List<Pessoa> lst = new List<Pessoa>();
            lst.Add(new Pessoa() { Id = 1, Name = "Jairo" });
            lst.Add(new Pessoa() { Id = 2, Name = "Leo" });
            lst.Add(new Pessoa() { Id = 3, Name = "Tinga" });

            return;

            //do
            //{
            //    Console.Clear();
            //    string op = "";

            //    Console.WriteLine("1 - inserir rodão");
            //    Console.WriteLine("2 - sair");
            //    op = Console.ReadLine();

            //    if (op.ToLower() != "s" && op.ToLower() != "n")
            //    {
            //        Console.ForegroundColor = ConsoleColor.Red;
            //        throw new JaderInputException($"o fei... Sei lá, só digitei {opcerteza} e ficou tudo vermelho aqui");
            //    }

            //    switch (op)
            //    {
            //        case "1":
            //            Console.WriteLine("digite o tamanho da roda para inserir no estoque");
            //            int tamanhoRoda = int.Parse(Console.ReadLine());
            //            if (tamanhoRoda < 22)
            //            {
            //                Console.ForegroundColor = ConsoleColor.Red;
            //                throw new JaderException("capai fei isso é roda de bicicleta");
            //            }
            //            else
            //            {
            //                listDeRodoes.Add(tamanhoRoda);
            //                Console.WriteLine("rodao masseta adicionado no estoque");
            //                Console.WriteLine();
            //                Console.WriteLine($"total de rodões: {listDeRodoes.Count}");
            //                Console.WriteLine("digite qualquer tecla para continuar");
            //                Console.ReadKey();
            //            }
            //            break;
            //        case "2":
            //            Console.WriteLine("nao vai mais umera mesmo? S/N");
            //            opcerteza = Console.ReadLine();
            //            if (opcerteza.ToLower() != "s" && opcerteza.ToLower() != "n")
            //            {
            //                Console.ForegroundColor = ConsoleColor.Red;
            //                throw new JaderInputException($"o fei... Sei lá, só digitei {opcerteza} e ficou tudo vermelho aqui");
            //            }
            //            break;
            //        default:
            //            Console.ForegroundColor = ConsoleColor.Red;
            //            throw new JaderInputException($"o fei... não sei oqq aconteceu aqui, mas tu digitou \"{op}\" e parou o programa");
            //    }

            //} while (opcerteza.ToLower() != "s");
        }
    }
}
