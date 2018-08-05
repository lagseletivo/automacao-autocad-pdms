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
    public class TubulacaoDrenagem
    {
        private static List<AtributosDoBloco> _lista;
        private static List<TextoElevacao> _listaElevacao;

        public TubulacaoDrenagem()
        {
            LerTodosOsBlocosEBuscarOsAtributos();
        }

        public static void LerTodosOsBlocosEBuscarOsAtributos()
        {
            Document documentoAtivo = Application.DocumentManager.MdiActiveDocument;
            Database database = documentoAtivo.Database;
            Editor editor = documentoAtivo.Editor;

            PromptSelectionResult pmtSelRes = editor.GetSelection();

            using (Transaction acTrans = documentoAtivo.TransactionManager.StartTransaction())
            {
                BlockTable blockTable;
                blockTable = acTrans.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;

                if (pmtSelRes.Status == PromptStatus.OK)
                {
                    _lista = new List<AtributosDoBloco>();
                    _listaElevacao = new List<TextoElevacao>();

                    foreach (ObjectId id in pmtSelRes.Value.GetObjectIds())
                    {
                        try
                        {                            

                            foreach (string nome in ConstantesTubulacao.TubulacaoNomeDosBlocos)
                            {
                                BlockTableRecord blockTableRecord;
                                blockTableRecord = acTrans.GetObject(blockTable[nome], OpenMode.ForRead) as BlockTableRecord;

                                if (!blockTableRecord.IsDynamicBlock) return;

                                try
                                {
                                    BlockReference bloco = null;

                                    bloco = acTrans.GetObject(id, OpenMode.ForRead) as BlockReference;

                                    DynamicBlockReferencePropertyCollection properties = bloco.DynamicBlockReferencePropertyCollection;

                                    BlockTableRecord nomeRealBloco = null;

                                    nomeRealBloco = acTrans.GetObject(bloco.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                                    if (blockTableRecord.Name == nome)
                                    {
                                        AtributosDoBloco atributo = BuscaPropriedadesDoBloco(bloco, properties, nomeRealBloco);
                                        BuscaDiametroDoTubo(atributo);

                                        double dimensaoFinalX;
                                        double dimensaoFinalY;

                                        //if (bloco.Rotation < 1.5708)
                                        //{
                                            dimensaoFinalX = atributo.Distancia * Math.Cos(bloco.Rotation);
                                            dimensaoFinalY = atributo.Distancia * Math.Sin(bloco.Rotation);

                                            atributo.XTubo = bloco.Position.X + dimensaoFinalX;
                                            atributo.YTubo = bloco.Position.Y + dimensaoFinalY;

                                        //}
                                        //else if (bloco.Rotation > 1.5708 && bloco.Rotation <= 3.14159265)
                                        //{
                                        //    dimensaoFinalX = distancia * Math.Sin(3.14159265 - bloco.Rotation);
                                        //    dimensaoFinalY = distancia * Math.Cos(3.14159265 - bloco.Rotation);

                                        //    atributo.XTubo = bloco.Position.X - dimensaoFinalX;
                                        //    atributo.YTubo = bloco.Position.Y + dimensaoFinalY;

                                        //}
                                        //else if (bloco.Rotation > 3.14159265 && bloco.Rotation <= 4.71239)
                                        //{
                                        //    dimensaoFinalX = distancia * Math.Sin(4.71239 - bloco.Rotation);
                                        //    dimensaoFinalY = distancia * Math.Cos(4.71239 - bloco.Rotation);

                                        //    atributo.XTubo = bloco.Position.X - dimensaoFinalX;
                                        //    atributo.YTubo = bloco.Position.Y - dimensaoFinalY;

                                        //}
                                        //else if (bloco.Rotation > 4.71239 && bloco.Rotation <= 6.28319)
                                        //{
                                        //    dimensaoFinalX = distancia * Math.Sin(6.28319 - bloco.Rotation);
                                        //    dimensaoFinalY = distancia * Math.Cos(6.28319 - bloco.Rotation);

                                        //    atributo.XTubo = bloco.Position.X + dimensaoFinalX;
                                        //    atributo.YTubo = bloco.Position.Y - dimensaoFinalY;

                                        //}
                                        _lista.Add(atributo);

                                        //------------------------------------------------------------------------------------------------------------
                                        TextoElevacao Elevacao1 = new TextoElevacao();

                                        Point3dCollection pntCol = new Point3dCollection
                                            {new Point3d(atributo.X - 5, atributo.Y + 5, 0),
                                             new Point3d(atributo.X + 5, atributo.Y + 5, 0),
                                             new Point3d(atributo.X + 5, atributo.Y - 5, 0),
                                             new Point3d(atributo.X - 5, atributo.Y - 5, 0)};

                                        PromptSelectionResult pmtSelResPoint = editor.SelectCrossingPolygon(pntCol);

                                        if (pmtSelResPoint.Status == PromptStatus.OK)
                                        {
                                            MText itemSelecionado = null;
                                            double distanciaMinima = Double.MaxValue;

                                            foreach (ObjectId objId in pmtSelResPoint.Value.GetObjectIds())
                                            {
                                                if (objId.ObjectClass.DxfName == "MTEXT")
                                                {
                                                    var text = acTrans.GetObject(objId, OpenMode.ForWrite) as MText;
                                                    if (text.Text.Contains("CF="))
                                                    {
                                                        double distanciaTexto = Math.Sqrt(Math.Pow(text.Location.X - atributo.X, 2) + Math.Pow(text.Location.Y - atributo.Y, 2));
                                                        if (distanciaTexto < distanciaMinima)
                                                        {
                                                            distanciaMinima = distanciaTexto;
                                                            itemSelecionado = text;
                                                        }
                                                    }
                                                }
                                            }
                                            if (itemSelecionado != null)
                                            {
                                                var listaTexto = itemSelecionado.Text.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                                                string textoCA = listaTexto.Where(d => d.Contains("CF=")).Any() ? listaTexto.Where(p => p.Contains("CF=")).FirstOrDefault() : string.Empty;
                                                textoCA = textoCA.Replace("CF=", "");

                                                Elevacao1.ElevacaoInicial = textoCA;
                                                Elevacao1.PosicaoX = itemSelecionado.Location.X;
                                                Elevacao1.PosicaoY = itemSelecionado.Location.Y;

                                            }

                                            else
                                                editor.WriteMessage("\nAs Elevações não foram encontradas!");
                                        }
                                        //-------------------------------------------------------------------------------------------------------------
                                        Point3dCollection pntColElevacao = new Point3dCollection
                                            {new Point3d(atributo.XTubo - 5, atributo.YTubo + 5, 0),
                                             new Point3d(atributo.XTubo + 5, atributo.YTubo + 5, 0),
                                             new Point3d(atributo.XTubo + 5, atributo.YTubo - 5, 0),
                                             new Point3d(atributo.XTubo - 5, atributo.YTubo - 5, 0)};

                                        PromptSelectionResult pmtSelResPoint2 = editor.SelectCrossingPolygon(pntColElevacao);

                                        if (pmtSelResPoint2.Status == PromptStatus.OK)
                                        {
                                            MText itemSelecionado2 = null;
                                            double distanciaMinima2 = Double.MaxValue;

                                            foreach (ObjectId objId2 in pmtSelResPoint2.Value.GetObjectIds())
                                            {
                                                if (objId2.ObjectClass.DxfName == "MTEXT")
                                                {
                                                    var text2 = acTrans.GetObject(objId2, OpenMode.ForWrite) as MText;
                                                    if (text2.Text.Contains("CF="))
                                                    {
                                                        double distanciaTexto2 = Math.Sqrt(Math.Pow(text2.Location.X - atributo.XTubo, 2) + Math.Pow(text2.Location.Y - atributo.YTubo, 2));
                                                        if (distanciaTexto2 < distanciaMinima2)
                                                        {
                                                            distanciaMinima2 = distanciaTexto2;
                                                            itemSelecionado2 = text2;
                                                        }
                                                    }
                                                }
                                            }
                                            if (itemSelecionado2 != null)
                                            {
                                                var listaTexto2 = itemSelecionado2.Text.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                                                string textoCA2 = listaTexto2.Where(d => d.Contains("CF=")).Any() ? listaTexto2.Where(p => p.Contains("CF=")).FirstOrDefault() : string.Empty;
                                                textoCA2 = textoCA2.Replace("CF=", "");

                                                Elevacao1.ElevacaoFinal = textoCA2;
                                                Elevacao1.PosicaoX = itemSelecionado2.Location.X;
                                                Elevacao1.PosicaoY = itemSelecionado2.Location.Y;

                                            }

                                            else
                                                editor.WriteMessage("\nAs Elevações não foram encontradas!");

                                            _listaElevacao.Add(Elevacao1);
                                        }
                                        //------------------------------------------------------------------------------------------------------------

                                        break;
                                    }

                                }
                                catch (Exception)
                                {
                                    continue;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                    }
                }
                acTrans.Commit();
            }
        }

        private static void BuscaDiametroDoTubo(AtributosDoBloco atributo)
        {
            var lista = atributo.NomeBloco.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (lista.Where(c => c.Contains("TUBO FF DN ")).Any())
            {
                string diametro = lista.Where(p => p.Contains("TUBO FF DN ")).Any() ? lista.Where(p => p.Contains("TUBO FF DN ")).FirstOrDefault() : string.Empty;
                diametro = diametro.Replace("TUBO FF DN ", "");
                atributo.Diametro = diametro;
            }
            else if (lista.Where(c => c.Contains("TUBO CONC ARMADO DN")).Any())
            {
                string diametro = lista.Where(p => p.Contains("TUBO CONC ARMADO DN")).Any() ? lista.Where(p => p.Contains("TUBO CONC ARMADO DN")).FirstOrDefault() : string.Empty;
                diametro = diametro.Replace("TUBO CONC ARMADO DN", "");
                atributo.Diametro = diametro;
            }
        }

        private static AtributosDoBloco BuscaPropriedadesDoBloco(BlockReference bloco, DynamicBlockReferencePropertyCollection properties, BlockTableRecord nomeRealBloco)
        {
            AtributosDoBloco atributo = new AtributosDoBloco();
            foreach (DynamicBlockReferenceProperty property in properties)
            {
                if (property.PropertyName == "Distance1")
                {
                    atributo.Distancia = Convert.ToDouble(property.Value);
                    break;
                }
            }
            atributo.X = bloco.Position.X;
            atributo.Y = bloco.Position.Y;
            atributo.NomeBloco = nomeRealBloco.Name;
            atributo.Handle = bloco.Handle.ToString();
            atributo.Angulo = bloco.Rotation;
            return atributo;
        }

        private static void FinalizaTarefasAposExcecao(string mensagemInicial, Exception excecao)
        {
            Console.WriteLine();
            Console.WriteLine(mensagemInicial + " Erro:" + Environment.NewLine + excecao.Message + Environment.NewLine + Environment.NewLine);
            Console.WriteLine("Pressione qualquer tecla para sair.");
            Environment.Exit(0);
        }

        public static void EscreveDadosNoExcel()
        {
            ExcelUtilsTubulacao.AbrirExcel();
            ExcelUtilsTubulacao.EscreveDados(_lista);
            ExcelUtilsTubulacao.EscreveElevacao(_listaElevacao);
        }
    }
}