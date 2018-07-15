using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Linq;
using System.Collections.Generic;
using AutocadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Automacao.Setup;
using System;

namespace Automacao
{
    public class SelecaoDosBlocos
    {
        private static AtributosDoBloco Atributo1 = new AtributosDoBloco();
        private static List<AtributosDoBloco> _lista;

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
            try
            {
                _lista = new List<AtributosDoBloco>();

                TypedValue[] filtro = { new TypedValue(0, "INSERT"), new TypedValue(0, typeof(BlockReference)) };
                PromptSelectionResult selecao = AutoCadDocument.Editor.SelectAll(new SelectionFilter(filtro));
                foreach (BlockReference bloco in selecao.Value)
                {
                    if (Constantes.PrefixoDoNomeDosBlocos.Contains(bloco.BlockName))
                    {
                        Atributo1.X = bloco.Position.X;
                        Atributo1.Y = bloco.Position.Y;
                        Atributo1.Handle = bloco.Handle.ToString();
                        Atributo1.NomeEfetivoDoBloco = bloco.BlockName;
                        Atributo1.Angulo = bloco.Rotation;
                        _lista.Add(Atributo1);
                    }
                }
            }

            catch (Exception e)
            {
                FinalizaTarefasAposExcecao("Ocorreu um erro ao ler os blocos do AutoCAD.", e);
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

        private static void FinalizaTarefasAposExcecao(string mensagemInicial, Exception excecao)
        {
            Console.WriteLine();
            Console.WriteLine(mensagemInicial + " Erro:" + Environment.NewLine + excecao.Message + Environment.NewLine + Environment.NewLine);
            Console.WriteLine("Pressione qualquer tecla para sair.");
            Console.ReadKey();
            Environment.Exit(0);
        }
    }
}

