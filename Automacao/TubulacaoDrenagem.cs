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
                            AtributosDoBloco Atributo1 = new AtributosDoBloco();

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

                                    if (blockTableRecord.IsDynamicBlock && blockTableRecord.Name == nome)
                                    {
                                        for (int i = 0; i < properties.Count; i++)
                                        {
                                            DynamicBlockReferenceProperty property = properties[i];

                                            if (property.PropertyName == "Distance1")
                                            {
                                                Atributo1.Distancia = property.Value.ToString();
                                                
                                            }
                                            Atributo1.X = bloco.Position.X;
                                            Atributo1.Y = bloco.Position.Y;
                                            Atributo1.NomeBloco = nomeRealBloco.Name;
                                            Atributo1.Handle = bloco.Handle.ToString();
                                            Atributo1.Angulo = bloco.Rotation;
                                            

                                            var lista = Atributo1.NomeBloco.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                                            if (lista.Where(c => c.Contains("TUBO FF DN ")).Any())
                                            {
                                                string diametro = lista.Where(p => p.Contains("TUBO FF DN ")).Any() ? lista.Where(p => p.Contains("TUBO FF DN ")).FirstOrDefault() : string.Empty;
                                                diametro = diametro.Replace("TUBO FF DN ", "");
                                                Atributo1.Diametro = diametro;
                                            }
                                            else if (lista.Where(c => c.Contains("TUBO CONC ARMADO DN")).Any())
                                            {
                                                string diametro = lista.Where(p => p.Contains("TUBO CONC ARMADO DN")).Any() ? lista.Where(p => p.Contains("TUBO CONC ARMADO DN")).FirstOrDefault() : string.Empty;
                                                diametro = diametro.Replace("TUBO CONC ARMADO DN", "");
                                                Atributo1.Diametro = diametro;
                                            }

                                            double distancia = Convert.ToDouble(Atributo1.Distancia);

                                            double dimensaoFinalX;
                                            double dimensaoFinalY;

                                            if (bloco.Rotation < 1.5708)
                                            {
                                                dimensaoFinalX = distancia * Math.Cos(bloco.Rotation);
                                                dimensaoFinalY = distancia * Math.Sin(bloco.Rotation);

                                                Atributo1.XTubo = bloco.Position.X + dimensaoFinalX;
                                                Atributo1.YTubo = bloco.Position.Y + dimensaoFinalY;
                                               
                                            }
                                            else if (bloco.Rotation > 1.5708 && bloco.Rotation <= 3.14159265)
                                            {
                                                dimensaoFinalX = distancia * Math.Sin(3.14159265 - bloco.Rotation);
                                                dimensaoFinalY = distancia * Math.Cos(3.14159265 - bloco.Rotation);

                                                Atributo1.XTubo = bloco.Position.X - dimensaoFinalX;
                                                Atributo1.YTubo = bloco.Position.Y + dimensaoFinalY;
                                               
                                            }
                                            else if (bloco.Rotation > 3.14159265 && bloco.Rotation <= 4.71239)
                                            {
                                                dimensaoFinalX = distancia * Math.Sin(4.71239 - bloco.Rotation);
                                                dimensaoFinalY = distancia * Math.Cos(4.71239 - bloco.Rotation);

                                                Atributo1.XTubo = bloco.Position.X - dimensaoFinalX;
                                                Atributo1.YTubo = bloco.Position.Y - dimensaoFinalY;
                                                
                                            }
                                            else if (bloco.Rotation > 4.71239 && bloco.Rotation <= 6.28319)
                                            {
                                                dimensaoFinalX = distancia * Math.Sin(6.28319 - bloco.Rotation);
                                                dimensaoFinalY = distancia * Math.Cos(6.28319 - bloco.Rotation);

                                                Atributo1.XTubo = bloco.Position.X + dimensaoFinalX;
                                                Atributo1.YTubo = bloco.Position.Y - dimensaoFinalY;
                                                
                                            }
                                            _lista.Add(Atributo1);

                                            //------------------------------------------------------------------------------------------------------------
                                            TextoElevacao Elevacao1 = new TextoElevacao();

                                            Point3dCollection pntCol = new Point3dCollection
                                            {new Point3d(Atributo1.X - 5, Atributo1.Y + 5, 0),
                                             new Point3d(Atributo1.X + 5, Atributo1.Y + 5, 0),
                                             new Point3d(Atributo1.X + 5, Atributo1.Y - 5, 0),
                                             new Point3d(Atributo1.X - 5, Atributo1.Y - 5, 0)};

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
                                                            double distanciaTexto = Math.Sqrt(Math.Pow(text.Location.X - Atributo1.X, 2) + Math.Pow(text.Location.Y - Atributo1.Y, 2));
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
                                            {new Point3d(Atributo1.XTubo - 5, Atributo1.YTubo + 5, 0),
                                             new Point3d(Atributo1.XTubo + 5, Atributo1.YTubo + 5, 0),
                                             new Point3d(Atributo1.XTubo + 5, Atributo1.YTubo - 5, 0),
                                             new Point3d(Atributo1.XTubo - 5, Atributo1.YTubo - 5, 0)};

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
                                                            double distanciaTexto2 = Math.Sqrt(Math.Pow(text2.Location.X - Atributo1.XTubo, 2) + Math.Pow(text2.Location.Y - Atributo1.YTubo, 2));
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