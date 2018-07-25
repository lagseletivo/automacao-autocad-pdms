using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Drenagem.Setup;
using System;
using System.Collections.Generic;
using System.Linq;
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
                                bloco = (BlockReference)acTrans.GetObject(objId_loopVariable, OpenMode.ForRead) as BlockReference;

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
                                Atributo1.NomeBloco = nomeRealBloco.Name;
                                Atributo1.Handle = bloco.Handle.ToString();
                                Atributo1.Angulo = bloco.Rotation;
                                _lista.Add(Atributo1);
                                //------------------------------------------------------------------------------------------------------------------------------------------------------
                                //------------------------------------------------------------------------------------------------------------------------------------------------------
                                TextoElevacao Elevacao1 = new TextoElevacao();

                                Point3dCollection pntCol = new Point3dCollection
                                {
                                    new Point3d(Atributo1.X - 7.5, Atributo1.Y + 7.5, 0),
                                    new Point3d(Atributo1.X + 7.5, Atributo1.Y + 7.5, 0),
                                    new Point3d(Atributo1.X + 7.5, Atributo1.Y - 7.5, 0),
                                    new Point3d(Atributo1.X - 7.5, Atributo1.Y - 7.5, 0)
                                };

                                PromptSelectionResult pmtSelRes = editor.SelectCrossingPolygon(pntCol);

                                if (pmtSelRes.Status == PromptStatus.OK)
                                {
                                    MText itemSelecionado = null;
                                    double distanciaMinima = Double.MaxValue;

                                    foreach (ObjectId id in pmtSelRes.Value.GetObjectIds())
                                    { 
                                        if (id.ObjectClass.DxfName == "MTEXT")
                                        {
                                            var text = acTrans.GetObject(id, OpenMode.ForWrite) as MText;
                                            if (text.Text.Contains("CA="))
                                            {
                                                double distancia = Math.Sqrt(Math.Pow(text.Location.X - Atributo1.X, 2) + Math.Pow(text.Location.Y - Atributo1.Y, 2));
                                                if (distancia < distanciaMinima)
                                                {
                                                    distanciaMinima = distancia;
                                                    itemSelecionado = text;
                                                }
                                            }                                           
                                        }
                                    }

                                    if (itemSelecionado != null)
                                    {
                                        var lista = itemSelecionado.Text.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                                        string textoCA = lista.Where(p => p.Contains("CA=")).Any() ? lista.Where(p => p.Contains("CA=")).FirstOrDefault() : string.Empty;
                                        string textoCFC = lista.Where(p => p.Contains("CFC=")).Any() ? lista.Where(p => p.Contains("CFC=")).FirstOrDefault() : string.Empty;

                                        textoCA = textoCA.Replace("CA=", "");
                                        textoCFC = textoCFC.Replace("CFC=", "");

                                        Elevacao1.ElevacaoInicial = textoCA;
                                        Elevacao1.ElevacaoFinal = textoCFC;
                                        Elevacao1.PosicaoX = itemSelecionado.Location.X;
                                        Elevacao1.PosicaoY = itemSelecionado.Location.Y;
                                        _listaElevacao.Add(Elevacao1);

                                    }

                                    else
                                        editor.WriteMessage("\nDid As Elevações não foram encontradas!");
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

