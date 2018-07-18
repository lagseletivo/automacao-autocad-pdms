using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automacao
{
    public static class ExcelUtils
    {    
        private static Excel.Worksheet _planilha;
        private static Excel.Workbook _pasta;

        public static Excel.Worksheet Planilha =>_planilha;
        public static Excel.Workbook Pasta =>_pasta;

        public static void AbrirExcel()
        {
            string nomeDoArquivo = AbreDialogoParaUsuarioSelecionarArquivo();
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            _pasta = excelApp.Workbooks.Open(nomeDoArquivo);
            _planilha = _pasta.ActiveSheet;        
        }
        
        public static void EscreveDados(List<AtributosDoBloco> lista)
        {
            int linha = 4;

            foreach (AtributosDoBloco atributo in lista)
            {
                _planilha.Cells[linha, 6] = atributo.X;
                _planilha.Cells[linha, 7] = atributo.Y;
                _planilha.Cells[linha, 1] = atributo.Handle;
                //planilha.Cells[linha, 2] = atributo.Nome;
                _planilha.Cells[linha, 3] = atributo.NomeEfetivoDoBloco;
                _planilha.Cells[linha, 25] = atributo.Angulo;
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