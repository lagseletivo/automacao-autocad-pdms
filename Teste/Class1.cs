using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutocadApp = Autodesk.AutoCAD.ApplicationServices.Application;


namespace Teste
{
    public class Class1
    {
        AtributosDoBloco Atributo1 = new AtributosDoBloco();

        List<String> PrefixoDoNomeDosBlocos = new List<String>
                {"ST","CA", "CPC","CP-OLEOSA",
                "CPNC",
                "CPO",
                "CPNC",
                "CL",
                "CPCD",
                "CCC"};

        public static Document AutoCadDocument
        {
            get { return AutocadApp.DocumentManager.MdiActiveDocument; }
        }

        public static SelectionSet GetSelectionSet()
        {
            var _editor = AutoCadDocument.Editor;
            var _selAll = _editor.SelectAll();
            return _selAll.Value;
        }

        public List<AtributosDoBloco> LerTodosOsBlocosEBuscarOsAtributos()
        {
            List<AtributosDoBloco> lista = new List<AtributosDoBloco>();

            SelectionSet selecao = GetSelectionSet();

            foreach (BlockReference bloco in selecao)
            {
                Atributo1.X = bloco.Position.X;
                Atributo1.Y = bloco.Position.Y;
                Atributo1.Handle = bloco.Handle.ToString();
                Atributo1.NomeEfetivoDoBloco = bloco.BlockName;
                Atributo1.Angulo = bloco.Rotation;
                lista.Add(Atributo1);
            }
            return lista;
        }

        public void EscreveDadosNoExcel(List<AtributosDoBloco> lista)
        {
            int linha = 4;
            foreach(AtributosDoBloco atributo in lista)
            {
                Excel.cells(linha, "F") = atributo.X;
                Excel.cells(linha, "G") = atributo.Y;
                Excel.cells(linha, "A") = atributo.Handle;
                Excel.cells(linha, "B") = atributo.NomeEfetivoDoBloco;
                Excel.cells(linha, "S") = atributo.Angulo;
                linha++;
            }
        }
    }
}
