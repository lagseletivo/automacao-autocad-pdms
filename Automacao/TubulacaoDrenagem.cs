using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Internal.DatabaseServices;
using Drenagem.Setup;
using System;
using System.Collections.Generic;
using System.Linq;
using TodosBlocos;

namespace Drenagem
{
    public class TubulacaoDrenagem
    {
        private static List<AtributosDoBloco> _lista;
        //private static List<TextoElevacao> _listaElevacao;

        public TubulacaoDrenagem()
        {
            LerTodosOsBlocosEBuscarOsAtributos();
        }

        public static void LerTodosOsBlocosEBuscarOsAtributos()
        {
            Document documentoAtivo = Application.DocumentManager.MdiActiveDocument;
            Database database = documentoAtivo.Database;
            Editor editor = documentoAtivo.Editor;

            using (Transaction acTrans = database.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = (BlockTable)acTrans.GetObject(database.BlockTableId, OpenMode.ForRead);

                try
                {
                    //_listaElevacao = new List<TextoElevacao>();

                    foreach (string nome in ConstantesTubulacao.TubulacaoNomeDosBlocos)
                    {
                        try
                        {
                            _lista = new List<AtributosDoBloco>();
 
                            foreach (ObjectId objId_loopVariable in blockTable)
                            {
                                BlockTableRecord blockTableRecord = (BlockTableRecord)acTrans.GetObject(objId_loopVariable, OpenMode.ForRead) as BlockTableRecord;
                                AtributosDoBloco Atributo1 = new AtributosDoBloco();

                                if (blockTableRecord.IsDynamicBlock && blockTableRecord.Name.Equals(nome))
                                {
                                    BlockReference bloco;
                                    bloco = (BlockReference)acTrans.GetObject(objId_loopVariable, OpenMode.ForRead) as BlockReference;
                                                                       
                                    //-------------------------------------------------------------------------------------------------------------------------------------
                                    //-------------------------------------------------------------------------------------------------------------------------------------
                                    //-------------------------------------------------------------------------------------------------------------------------------------

                                    BlockTableRecord btrDyn = acTrans.GetObject(bloco.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                    if (btrDyn != null && !btrDyn.ExtensionDictionary.IsNull)
                                    {
                                        DBDictionary extDic = acTrans.GetObject(btrDyn.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                                        if (extDic != null)
                                        {
                                            EvalGraph graph = acTrans.GetObject(extDic.GetAt("ACAD_ENHANCEDBLOCK"), OpenMode.ForRead) as EvalGraph;

                                            int[] nodeIds = graph.GetAllNodes();
                                            foreach (uint nodeId in nodeIds)
                                            {
                                                DBObject node = graph.GetNode(nodeId, OpenMode.ForRead, acTrans);
                                                if (!(node is BlockPropertiesTable)) continue;
                                                BlockPropertiesTable bpt = node as BlockPropertiesTable;
                                                int currentRow = SelectRowNumber(ref bpt);
                                                BlockPropertiesTableRow row = bpt.Rows[currentRow];
                                                List<string> nameProps = new List<string>();
                                                for (int i = 0; i < bpt.Columns.Count; i++)
                                                {
                                                    nameProps.Add(bpt.Columns[i].Parameter.Name);
                                                }
                                                DynamicBlockReferencePropertyCollection dynPropsCol = bloco.DynamicBlockReferencePropertyCollection;

                                                foreach (DynamicBlockReferenceProperty dynProp in dynPropsCol)
                                                {
                                                    int i = nameProps.FindIndex(delegate (string s) { return s == dynProp.PropertyName; });
                                                    if (i >= 0 && i < nameProps.Count)
                                                    {
                                                        try
                                                        {
                                                            dynProp.Value = row[i].AsArray()[0].Value;
                                                        }
                                                        catch
                                                        {
                                                            editor.WriteMessage("\nCan not set to {0} value={1}",
                                                              dynProp.PropertyName, row[i].AsArray()[0].Value);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }


                                    //-------------------------------------------------------------------------------------------------------------------------------------
                                    //-------------------------------------------------------------------------------------------------------------------------------------
                                    //-------------------------------------------------------------------------------------------------------------------------------------

                                }
                            }
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                        continue;
                    }
                }

                catch (Exception e)
                {
                    FinalizaTarefasAposExcecao("Ocorreu um erro ao ler os blocos do AutoCAD.", e);
                }
                acTrans.Commit();
            }
        }
        public static int SelectRowNumber(ref BlockPropertiesTable bpt)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            int columns = bpt.Columns.Count;
            int rows = bpt.Rows.Count;
            int currentRow = 0, currentColumn = 0;
            ed.WriteMessage("\n");
            for (currentColumn = 0; currentColumn < columns; currentColumn++)
            {
                ed.WriteMessage("{0}; ", bpt.Columns[currentColumn].Parameter.Name);
            }
            foreach (BlockPropertiesTableRow row in bpt.Rows)
            {
                ed.WriteMessage("\n[{0}]:\t", currentRow);
                for (currentColumn = 0; currentColumn < columns; currentColumn++)
                {
                    TypedValue[] columnValue = row[currentColumn].AsArray();
                    foreach (TypedValue tpVal in columnValue)
                    {
                        ed.WriteMessage("{0}; ", tpVal.Value);
                    }
                    ed.WriteMessage("|");
                }
                currentRow++;
            }

            PromptIntegerResult res;
            string.Format("0-{0}", rows - 1);

            while ((res = ed.GetInteger(string.Format("\nSelect row number (0-{0}): ", rows - 1))).Status == PromptStatus.OK)
            {
                if (res.Value >= 0 && res.Value <= rows) return res.Value;
            }
            return -1;
        }
        //public static void EscreveDadosNoExcel()
        //{
        //    ExcelUtils.AbrirExcel();
        //    ExcelUtils.EscreveDados(_lista);
        //    //ExcelUtils.EscreveElevacao(_listaElevacao);

        //}
        private static void FinalizaTarefasAposExcecao(string mensagemInicial, Exception excecao)
        {
            Console.WriteLine();
            Console.WriteLine(mensagemInicial + " Erro:" + Environment.NewLine + excecao.Message + Environment.NewLine + Environment.NewLine);
            Console.WriteLine("Pressione qualquer tecla para sair.");
            Environment.Exit(0);
        }


    }
}

