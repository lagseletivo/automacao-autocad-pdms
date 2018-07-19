using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Internal.DatabaseServices;
using Drenagem.Setup;
using System;
using System.Collections.Generic;
using TodosBlocos;


namespace Drenagem
{
    public class TubulacaoDrenagem
    {
        private static List<AtributosDoBloco> _lista;

        public TubulacaoDrenagem()
        {
            LerTodosOsBlocosEBuscarOsAtributos();
        }
        public static void LerTodosOsBlocosEBuscarOsAtributos()
        {
            Document documentoAtivo = Application.DocumentManager.MdiActiveDocument;
            Database database = documentoAtivo.Database;
            Editor editor = documentoAtivo.Editor;

            ObjectId idBTR = ObjectId.Null;

            using (Transaction acTrans = database.TransactionManager.StartTransaction())
            {
                using (BlockTable acBlockTable = database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable)
                {
                    BlockTable blockTable;
                    blockTable = acTrans.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;

                    try
                    {
                        _lista = new List<AtributosDoBloco>();

                        foreach (string nome in ConstantesTubulacao.TubulacaoNomeDosBlocos)
                        {
                            BlockTableRecord blockTableRecord;
                            blockTableRecord = acTrans.GetObject(blockTable[nome], OpenMode.ForRead) as BlockTableRecord;

                            BlockReference blocoRefDinamico;
                            blocoRefDinamico = (BlockReference)acTrans.GetObject(blockTable[nome], OpenMode.ForRead) as BlockReference;

                            BlockTableRecord blockTableRecordDynamic = acTrans.GetObject(blocoRefDinamico.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                            foreach (ObjectId objId_loopVariable in blockTableRecord.GetBlockReferenceIds(true, true))
                            {

                                if (blockTableRecordDynamic != null && !blockTableRecordDynamic.ExtensionDictionary.IsNull)
                                {
                                    DBDictionary extDic = acTrans.GetObject(blockTableRecordDynamic.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;

                                    if (extDic != null)
                                    {
                                        EvalGraph graph = acTrans.GetObject(extDic.GetAt("ACAD_ENHANCeditorBLOCK"), OpenMode.ForRead) as EvalGraph;
                                        int[] nodeIds = graph.GetAllNodes();

                                        foreach (uint nodeId in nodeIds)
                                        {
                                            DBObject node = graph.GetNode(nodeId, OpenMode.ForRead, acTrans);
                                            if (!(node is BlockPropertiesTable)) continue;
                                            BlockPropertiesTable blockPropertiesTable = node as BlockPropertiesTable;
                                            int currentRow = SelectRowNumber(ref blockPropertiesTable);
                                            BlockPropertiesTableRow row = blockPropertiesTable.Rows[currentRow];
                                            List<string> nameProps = new List<string>();

                                            for (int i = 0; i < blockPropertiesTable.Columns.Count; i++)
                                            {
                                                nameProps.Add(blockPropertiesTable.Columns[i].Parameter.Name);
                                            }
                                            DynamicBlockReferencePropertyCollection dynPropsCol = blocoRefDinamico.DynamicBlockReferencePropertyCollection;

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
                                                    //-----------------------------------------------------------------------------------------------
                                                    //foreach (ObjectId objId_loopVariable in blocoRefDinamico.GetBlockReferenceIds(true, true))
                                                    AtributosDoBloco Atributo1 = new AtributosDoBloco();

                                                    Atributo1.X = blocoRefDinamico.Position.X;
                                                    Atributo1.Y = blocoRefDinamico.Position.Y;
                                                    Atributo1.Handle = blocoRefDinamico.Handle.ToString();
                                                    Atributo1.Angulo = blocoRefDinamico.Rotation;

                                                    _lista.Add(Atributo1);

                                                }
                                            }
                                        }
                                    }
                                }
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
        }
        public static int SelectRowNumber(ref BlockPropertiesTable bpt)
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            int columns = bpt.Columns.Count;
            int rows = bpt.Rows.Count;
            int currentRow = 0, currentColumn = 0;
            editor.WriteMessage("\n");
            for (currentColumn = 0; currentColumn < columns; currentColumn++)
            {
                editor.WriteMessage("{0}; ", bpt.Columns[currentColumn].Parameter.Name);
            }
            foreach (BlockPropertiesTableRow row in bpt.Rows)
            {
                editor.WriteMessage("\n[{0}]:\t", currentRow);
                for (currentColumn = 0; currentColumn < columns; currentColumn++)
                {
                    TypedValue[] columnValue = row[currentColumn].AsArray();
                    foreach (TypedValue tpVal in columnValue)
                    {
                        editor.WriteMessage("{0}; ", tpVal.Value);
                    }
                    editor.WriteMessage("|");
                }
                currentRow++;
            }

            PromptIntegerResult res;
            string.Format("0-{0}", rows - 1);

            while ((res = editor.GetInteger(string.Format("\nSelect row number (0-{0}): ", rows - 1))).Status == PromptStatus.OK)
            {
                if (res.Value >= 0 && res.Value <= rows) return res.Value;
            }
            return -1;
        }

        public static void EscreveDadosNoExcel()
        {
            SelecaoDosTextosTubulacao.ZoomWindow();
            ExcelUtilsTubulacao.AbrirExcel();
            ExcelUtilsTubulacao.EscreveDados(_lista);
            //ExcelUtilsTubulacao.EscreveElevacao(_listaElevacao);
        }
        private static void FinalizaTarefasAposExcecao(string mensagemInicial, Exception excecao)
        {
            Console.WriteLine();
            Console.WriteLine(mensagemInicial + " Erro:" + Environment.NewLine + excecao.Message + Environment.NewLine + Environment.NewLine);
            Console.WriteLine("Pressione qualquer tecla para sair.");
            Environment.Exit(0);
        }
    }
}
