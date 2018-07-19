using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
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

            using (Transaction acTrans = database.TransactionManager.StartTransaction())
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

                        foreach (ObjectId objId_loopVariable in blockTableRecord)
                        {
                            BlockReference blocoDinamico;
                            blocoDinamico = (BlockReference)acTrans.GetObject(objId_loopVariable, OpenMode.ForRead) as BlockReference;

                            DynamicBlockReferencePropertyCollection properties = blocoDinamico.DynamicBlockReferencePropertyCollection;

                            AtributosDoBloco Atributo1 = new AtributosDoBloco();

                            for (int i = 0; i < properties.Count; i++)
                            {
                                DynamicBlockReferenceProperty property = properties[i];

                                if (property.PropertyName == "Distance1")
                                {
                                    Atributo1.Distancia = property.Value.ToString();
                                }
                            }
                            Atributo1.X = blocoDinamico.Position.X;
                            Atributo1.Y = blocoDinamico.Position.Y;
                            Atributo1.nomeBloco = blocoDinamico.Name;
                            Atributo1.Handle = blocoDinamico.Handle.ToString();
                            Atributo1.Angulo = blocoDinamico.Rotation;

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

