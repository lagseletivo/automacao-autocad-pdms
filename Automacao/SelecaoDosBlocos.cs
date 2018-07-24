using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Drenagem.Setup;
using System;
using System.Collections.Generic;
using TodosBlocos;

namespace Drenagem
{
    public class SelecaoDosBlocos
    {
        private static List<AtributosDoBloco> _lista;
        private static List<TextoElevacao> _listaElevacao;

        public SelecaoDosBlocos()
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
                BlockTable blockTable;
                blockTable = acTrans.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;

                try
                {
                    _lista = new List<AtributosDoBloco>();
                    _listaElevacao = new List<TextoElevacao>();

                    foreach (string nome in Constantes.PrefixoDoNomeDosBlocos)
                    {
                        try
                        {
                            BlockTableRecord blockTableRecord;
                            blockTableRecord = acTrans.GetObject(blockTable[nome], OpenMode.ForRead) as BlockTableRecord;

                            foreach (ObjectId objId_loopVariable in blockTableRecord.GetBlockReferenceIds(true, true))
                            {
                                AtributosDoBloco Atributo1 = new AtributosDoBloco();

                                BlockReference bloco;
                                bloco = (BlockReference)acTrans.GetObject(objId_loopVariable, OpenMode.ForRead) as BlockReference; ;

                                BlockTableRecord nomeRealBloco = null;

                                nomeRealBloco = acTrans.GetObject(bloco.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                                AttributeCollection attCol = bloco.AttributeCollection;

                                foreach (ObjectId attId in attCol)
                                {
                                    AttributeReference attRef = (AttributeReference)acTrans.GetObject(attId, OpenMode.ForRead);
                                    string texto = (attRef.TextString);
                                    //string tag = attRef.Tag;
                                    Atributo1.NomeEfetivoDoBloco = texto;
                                }

                                Atributo1.X = bloco.Position.X;
                                Atributo1.Y = bloco.Position.Y;
                                Atributo1.nomeBloco = nomeRealBloco.Name;
                                Atributo1.Handle = bloco.Handle.ToString();
                                Atributo1.Angulo = bloco.Rotation;
                                _lista.Add(Atributo1);
                                //------------------------------------------------------------------------------------------------------------------------------------------------------
                                //------------------------------------------------------------------------------------------------------------------------------------------------------
                                TextoElevacao Elevacao1 = new TextoElevacao();

                                Point3d p1 = new Point3d(Atributo1.X - 7.5, Atributo1.Y + 7.5, 0);
                                Point3d p2 = new Point3d(Atributo1.X + 7.5, Atributo1.Y + 7.5, 0);
                                Point3d p3 = new Point3d(Atributo1.X + 7.5, Atributo1.Y - 7.5, 0);
                                Point3d p4 = new Point3d(Atributo1.X - 7.5, Atributo1.Y - 7.5, 0);

                                Point3dCollection pntCol = new Point3dCollection();
                                pntCol.Add(p1);
                                pntCol.Add(p2);
                                pntCol.Add(p3);
                                pntCol.Add(p4);

                                PromptSelectionResult pmtSelRes = null;

                                //ObjectIdCollection oPlTxtIdColl = null;

                                //var filterList = new TypedValue[1];
                                //filterList.SetValue(new TypedValue(0, "MTEXT"), 0);                                

                                //SelectionFilter selecaoFiltro = new SelectionFilter(filterList);

                                pmtSelRes = editor.SelectCrossingPolygon(pntCol);

                                //int textoEncontrado = 0;

                                if (pmtSelRes.Status == PromptStatus.OK)
                                {                           
                                    foreach (ObjectId id in pmtSelRes.Value.GetObjectIds())
                                    { 
                                        if (id.ObjectClass.DxfName == "MTEXT")
                                        {
                                            var text = acTrans.GetObject(id, OpenMode.ForWrite) as MText;

                                            //var text = (DBText)acTrans.GetObject(id, OpenMode.ForRead);
                                            
                                            if (text.Text.Contains("CA="))
                                            {
                                                //Point3d locacao = text.Location.TransformBy(Matrix3d.Identity);

                                                Elevacao1.ElevacaoInicial = text.Text;
                                                //Elevacao1.PosicaoX = text.Location;
                                                //Elevacao1.PosicaoX = locacao;
                                                _listaElevacao.Add(Elevacao1);
                                            }
                                            if (text.Text.Contains("CFC="))
                                            {
                                                Elevacao1.ElevacaoFinal = text.Text;
                                                _listaElevacao.Add(Elevacao1);
                                            }
                                            else
                                                editor.WriteMessage("\nDid As Elevações não foram encontradas!");

                                            //textoEncontrado++;
                                        }
                                    }
                                }
                                //------------------------------------------------------------------------------------------------------------------------------------------------------
                                //------------------------------------------------------------------------------------------------------------------------------------------------------
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
        public static void EscreveDadosNoExcel()
        {
            ExcelUtils.AbrirExcel();
            ExcelUtils.EscreveDados(_lista);
            ExcelUtils.EscreveElevacao(_listaElevacao);

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

