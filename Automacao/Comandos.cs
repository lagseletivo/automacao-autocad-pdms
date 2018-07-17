using Autodesk.AutoCAD.Runtime;

namespace Automacao
{
    public class Comandos
    {
        [CommandMethod("Drenagem")]
        public static void Drenagem()
        {
           SelecaoDosBlocos.LerTodosOsBlocosEBuscarOsAtributos();
           SelecaoDosBlocos.LerTodosOsBlocosEBuscarOsAtributos();
            SelecaoDosBlocos.EscreveDadosNoExcel();
        }
    }
}
