using Autodesk.AutoCAD.Runtime;

namespace Teste
{
    public class Comando
    {
        [CommandMethod("Drenagem")]
        public static void Drenagem()
        {
            SelecaoDosBlocos comando = new SelecaoDosBlocos();

            //var lista = SelecaoDosBlocos.LerTodosOsBlocosEBuscarOsAtributos();
            //SelecaoDosBlocos.EscreveDadosNoExcel(lista);
        }
    }
}
