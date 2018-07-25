using Autodesk.AutoCAD.Runtime;
using Drenagem;
using Eletrica;

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
            //TubulacaoDrenagem.EscreveDadosNoExcel();
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

        [CommandMethod("EletricaEstaqueamento")]
        public static void EletricaEstaqueamento()
        {
            EstaqueamentoEletrica.LerTodosOsBlocosEBuscarOsAtributos();
            EstaqueamentoEletrica.LerTodosOsBlocosEBuscarOsAtributos();
            EstaqueamentoEletrica.EscreveDadosNoExcel();
        }
    }


}
