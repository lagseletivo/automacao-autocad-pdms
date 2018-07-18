using Autodesk.AutoCAD.Runtime;
using Eletrica;
using Drenagem;

namespace TodosComandos
{
    public class Comandos
    {
        [CommandMethod("DrenagemCaixas")]
        public static void DrenagemCaixas()
        {
            SelecaoDosBlocos.LerTodosOsBlocosEBuscarOsAtributos();
            SelecaoDosBlocos.LerTodosOsBlocosEBuscarOsAtributos();
            SelecaoDosBlocos.EscreveDadosNoExcel();
        }
        [CommandMethod("DrenagemTubulacao")]
        public static void DrenagemTubulacao()
        {
            TubulacaoDrenagem.LerTodosOsBlocosEBuscarOsAtributos();
            TubulacaoDrenagem.LerTodosOsBlocosEBuscarOsAtributos();
            TubulacaoDrenagem.EscreveDadosNoExcel();
        }
    }
    public class ComandosEletrica
    {
        [CommandMethod("EletricaCaixas")]
        public static void EletricaCaixas()
        {
            CaixaPassagemEletrica.LerTodosOsBlocosEBuscarOsAtributos();
            CaixaPassagemEletrica.LerTodosOsBlocosEBuscarOsAtributos();
            CaixaPassagemEletrica.EscreveDadosNoExcel();
        }
    }

}
