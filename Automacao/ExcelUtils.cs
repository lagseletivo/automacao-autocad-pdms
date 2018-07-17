using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automacao
{
    public static class ExcelUtils
    {
        public static void AbrirExcel()
        {
            //string _planilha = @"C:\Users\felip\Documents\Modelagem\Planilha_modelo.xlsm";
            //var _excelApp = new Excel.Application();
            //_excelApp.Visible = true;

            //Excel.Workbooks books = _excelApp.Workbooks;
            //Excel.Workbook sheet = books.Open(_planilha);
            string nomeDoArquivo=AbreDialogoParaUsuarioSelecionarArquivo();
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook pasta = excelApp.Workbooks.Open(nomeDoArquivo);
            Excel.Worksheet planilha = pasta.ActiveSheet;
            planilha.Cells[4, 1] = "ffuncionaooooooooooooooooooi";


        }

        //static void Main(string[] args)
        //{
        //    AbrirExcel();
        //}


        private static string AbreDialogoParaUsuarioSelecionarArquivo()
        {
            OpenFileDialog dialogo = new OpenFileDialog()
            {
                Title = "Selecione o arquivo",
                CheckFileExists = true,
                InitialDirectory = "C:"
            };
            dialogo.ShowDialog();
            if (!(dialogo.SafeFileName.Trim()==null || dialogo.SafeFileName.Trim()==""))
            {
                Console.WriteLine("Arquivo {0} selecionado.", dialogo.SafeFileName);
                return dialogo.FileName;
            }
            else
            {
                Console.WriteLine("Nenhum arquivo selecionado. Pressione qualquer tecla para sair.");
                Console.ReadKey();
                Environment.Exit(0);
                return "";
            }
        }
    }
}