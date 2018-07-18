using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Drenagem.Setup;
using TodosBlocos;
using System;
using System.Collections.Generic;

namespace Drenagem
{
    public class SelecaoDosBlocos
    {
        private static List<AtributosDoBloco> _lista;

        public SelecaoDosBlocos()
        {
            LerTodosOsBlocosEBuscarOsAtributos();
        }
        public static void LerTodosOsBlocosEBuscarOsAtributos()
        {
            Document documentoAtivo = Application.DocumentManager.MdiActiveDocument;
            Database database = documentoAtivo.Database;

            using (Transaction acTrans = database.TransactionManager.StartTransaction())
            {
                BlockTable blockTable;
                blockTable = acTrans.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;

                try
                {
                    _lista = new List<AtributosDoBloco>();

                    foreach (string nome in Constantes.PrefixoDoNomeDosBlocos)
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
                                string tag = attRef.Tag;
                                Atributo1.NomeEfetivoDoBloco = texto;
                            }

                            Atributo1.X = bloco.Position.X;
                            Atributo1.Y = bloco.Position.Y;
                            Atributo1.nomeBloco = nomeRealBloco.Name;
                            Atributo1.Handle = bloco.Handle.ToString();
                            Atributo1.Angulo = bloco.Rotation;
                            _lista.Add(Atributo1);
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

