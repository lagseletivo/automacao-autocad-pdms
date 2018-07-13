using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using AutocadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace Teste
{
    public class SelecaoDosBlocos
    {
        //[CommandMethod("Drenagem")]

        AtributosDoBloco Atributo1 = new AtributosDoBloco();

        List<String> PrefixoDoNomeDosBlocos = new List<String>
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
                if (PrefixoDoNomeDosBlocos.Contains(Atributo1.NomeEfetivoDoBloco))
                {
                    Atributo1.X = bloco.Position.X;
                    Atributo1.Y = bloco.Position.Y;
                    Atributo1.Handle = bloco.Handle.ToString();
                    Atributo1.NomeEfetivoDoBloco = bloco.BlockName;
                    Atributo1.Angulo = bloco.Rotation;
                    lista.Add(Atributo1);
                }
            }
            return lista;
        }

        public void EscreveDadosNoExcel(List<AtributosDoBloco> lista)
        {
            ExcelUtils.abrirExcel();

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

