using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using AutocadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Linq;

namespace Teste
{
    public class SelecaoDosBlocos
    {
        private static AtributosDoBloco Atributo1 = new AtributosDoBloco();
        private static List<AtributosDoBloco> _lista;

        private static List<String> PrefixoDoNomeDosBlocos = new List<String>
            { "ST",
               "CA",
               "CPC",
               "CP-OLEOSA",
               "CPNC",
               "CPO",
               "CPNC",
               "CL",
               "CPCD",
               "CCC" };

        public SelecaoDosBlocos()
        {
            LerTodosOsBlocosEBuscarOsAtributos();
            AutoCadDocument.Editor.WriteMessage("olar");
            AutoCadDocument.Editor.WriteMessage("esses sao teus bloquin:");            
            
            foreach (AtributosDoBloco a in _lista)
                AutoCadDocument.Editor.WriteMessage(a.NomeEfetivoDoBloco);

            // EscreveDadosNoExcel();
        }

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

        public static void LerTodosOsBlocosEBuscarOsAtributos()
        {
            _lista = new List<AtributosDoBloco>();

            SelectionSet selecao = GetSelectionSet();

            foreach (var bloco in selecao)
            {
                if (bloco.GetType() == typeof(BlockReference) && PrefixoDoNomeDosBlocos.Contains(((BlockReference)bloco).BlockName))
                {
                    BlockReference blocoTemporario = (BlockReference)bloco;

                    Atributo1.X = blocoTemporario.Position.X;
                    Atributo1.Y = blocoTemporario.Position.Y;
                    Atributo1.Handle = blocoTemporario.Handle.ToString();
                    Atributo1.NomeEfetivoDoBloco = blocoTemporario.BlockName;
                    Atributo1.Angulo = blocoTemporario.Rotation;
                    _lista.Add(Atributo1);
                }
            }
        }

        public static void EscreveDadosNoExcel()
        {
            ExcelUtils.AbrirExcel();

            //int linha = 4;
            //foreach (AtributosDoBloco atributo in lista)
            //{
            //    excel.Cells[linha, "F"].Text = atributo.X;
            //    excel.Cells[linha, "G"].Text = atributo.Y;
            //    excel.Cells[linha, "A"].Text = atributo.Handle;
            //    excel.Cells[linha, "B"].Text = atributo.NomeEfetivoDoBloco;
            //    excel.Cells[linha, "S"].Text = atributo.Angulo;

            //    linha++;
        }

    }
}

