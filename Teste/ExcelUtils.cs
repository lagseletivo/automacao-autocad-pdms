using Excel = Microsoft.Office.Interop.Excel;

namespace Teste
{
    public static class ExcelUtils
    {
        public static void AbrirExcel()
        {
            string _planilha = @"C:\Users\felip\Documents\Modelagem\Planilha_modelo.xlsm";
            var _excelApp = new Excel.Application();
            _excelApp.Visible = true;

            Excel.Workbooks books = _excelApp.Workbooks;
            Excel.Workbook sheet = books.Open(_planilha);
        }

        //static void Main(string[] args)
        //{
        //    AbrirExcel();
        //}


    }
}