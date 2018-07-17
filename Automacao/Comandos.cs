using Autodesk.AutoCAD.Runtime;

namespace Automacao
{
    public class Comandos
    {
        [CommandMethod("Drenagem")]
        public static void Drenagem()
        {
            SelecaoDosBlocos comando = new SelecaoDosBlocos();
            //SelecaoDosBlocos.LerTodosOsBlocosEBuscarOsAtributos();
            SelecaoDosBlocos.EscreveDadosNoExcel();
        }
    }
}
