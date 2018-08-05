using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Drenagem.Setup;
using System;
using System.Collections.Generic;
using System.Linq;
using TodosBlocos;

namespace Drenagem
{
    public class TubulacaoConexoes
    {
        private static List<AtributosDoBloco> _lista;

        public static void LerTodosOsBlocosEBuscarOsAtributos()
        {
            Document documentoAtivo = Application.DocumentManager.MdiActiveDocument;
            Database database = documentoAtivo.Database;
            Editor editor = documentoAtivo.Editor;


            using (Transaction acTrans = documentoAtivo.TransactionManager.StartTransaction())
            {
                BlockTable blockTable;
                blockTable = acTrans.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;

                //TypedValue[] acTypValAr = new TypedValue[0];
                // acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "BLOCK"),0);

                //SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);

                PromptSelectionResult pmtSelRes = editor.GetSelection();

                if (pmtSelRes.Status == PromptStatus.OK)
                {
                    _lista = new List<AtributosDoBloco>();

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
                                    if (blockTableRecord.IsDynamicBlock && blockTableRecord.Name.Equals(nome))
                                    {
                                        BlockReference bloco = acTrans.GetObject(id, OpenMode.ForRead) as BlockReference;

                                        DynamicBlockReferencePropertyCollection properties = bloco.DynamicBlockReferencePropertyCollection;

                                        BlockTableRecord nomeRealBloco = null;

                                        nomeRealBloco = acTrans.GetObject(bloco.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                        
                                        if (!_lista.Any(b=> b.Handle.ToString() == Atributo1.Handle))
                                        {

                                            for (int i = 0; i < properties.Count; i++)
                                            {
                                                DynamicBlockReferenceProperty property = properties[i];
                                                if (property.PropertyName == "Distance1")
                                                {
                                                    Atributo1.Distancia = Convert.ToDouble(property.Value);
                                                    _lista.Add(Atributo1);
                                                }
                                                Atributo1.X = bloco.Position.X;
                                                Atributo1.Y = bloco.Position.Y;
                                                Atributo1.NomeBloco = nomeRealBloco.Name;
                                                Atributo1.Handle = bloco.Handle.ToString();
                                                Atributo1.Angulo = bloco.Rotation;
                                                _lista.Add(Atributo1);

                                                var lista = Atributo1.NomeBloco.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                                                string diametro = lista.Where(p => p.Contains("TUBO FF DN ")).Any() ? lista.Where(p => p.Contains("TUBO FF DN ")).FirstOrDefault() : string.Empty;

                                                diametro = diametro.Replace("TUBO FF DN ", "");

                                                Atributo1.Diametro = diametro;

                                                double distancia = Convert.ToDouble(Atributo1.Distancia);

                                                double dimensaoFinalX;
                                                double dimensaoFinalY;

                                                if (bloco.Rotation < 1.5708)
                                                {
                                                    dimensaoFinalX = distancia * Math.Cos(bloco.Rotation);
                                                    dimensaoFinalY = distancia * Math.Sin(bloco.Rotation);

                                                    Atributo1.XTubo = bloco.Position.X + dimensaoFinalX;
                                                    Atributo1.YTubo = bloco.Position.Y + dimensaoFinalY;
                                                    _lista.Add(Atributo1);
                                                }
                                                else if (bloco.Rotation > 1.5708 && bloco.Rotation <= 3.14159265)
                                                {
                                                    dimensaoFinalX = distancia * Math.Sin(3.14159265 - bloco.Rotation);
                                                    dimensaoFinalY = distancia * Math.Cos(3.14159265 - bloco.Rotation);

                                                    Atributo1.XTubo = bloco.Position.X - dimensaoFinalX;
                                                    Atributo1.YTubo = bloco.Position.Y + dimensaoFinalY;
                                                    _lista.Add(Atributo1);
                                                }
                                                else if (bloco.Rotation > 3.14159265 && bloco.Rotation <= 4.71239)
                                                {
                                                    dimensaoFinalX = distancia * Math.Sin(4.71239 - bloco.Rotation);
                                                    dimensaoFinalY = distancia * Math.Cos(4.71239 - bloco.Rotation);

                                                    Atributo1.XTubo = bloco.Position.X - dimensaoFinalX;
                                                    Atributo1.YTubo = bloco.Position.Y - dimensaoFinalY;
                                                    _lista.Add(Atributo1);
                                                }
                                                else if (bloco.Rotation > 4.71239 && bloco.Rotation <= 6.28319)
                                                {
                                                    dimensaoFinalX = distancia * Math.Sin(6.28319 - bloco.Rotation);
                                                    dimensaoFinalY = distancia * Math.Cos(6.28319 - bloco.Rotation);

                                                    Atributo1.XTubo = bloco.Position.X + dimensaoFinalX;
                                                    Atributo1.YTubo = bloco.Position.Y - dimensaoFinalY;
                                                    _lista.Add(Atributo1);
                                                }
                                            }

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
        }
    }
}