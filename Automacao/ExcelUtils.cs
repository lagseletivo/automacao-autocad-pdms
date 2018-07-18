using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automacao
{
    public static class ExcelUtils
    {
        public static void AbrirExcel()
        {

            string nomeDoArquivo = AbreDialogoParaUsuarioSelecionarArquivo();
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook pasta = excelApp.Workbooks.Open(nomeDoArquivo);
            Excel.Worksheet planilha = pasta.ActiveSheet;

            int linha = 4;

            foreach (AtributosDoBloco atributo in SelecaoDosBlocos._lista)
            {
                planilha.Cells[linha, 6] = atributo.X;
                planilha.Cells[linha, 7] = atributo.Y;
                planilha.Cells[linha, 1] = atributo.Handle;
                planilha.Cells[linha, 2] = atributo.nomeBloco;
                planilha.Cells[linha, 3] = atributo.NomeEfetivoDoBloco;
                planilha.Cells[linha, 25] = atributo.Angulo;

                linha++;
            }
          
        }

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