using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace Drenagem
{
    public class TubulacaoDrenagem
    {
        public static void GetPointsFromUser()
        {
            // Get the current database and start the Transaction Manager
            Document documentoAtivo = Application.DocumentManager.MdiActiveDocument;
            Database database = documentoAtivo.Database;
            Editor editor = documentoAtivo.Editor;

            PromptPointResult pPtRes;
            PromptPointOptions pPtOpts = new PromptPointOptions("");
            // Prompt for the start point
            pPtOpts.Message = "\nEnter the start point of the line: ";
            pPtRes = documentoAtivo.Editor.GetPoint(pPtOpts);
            Point3d ptStart = pPtRes.Value;// Exit if the user presses ESC or cancels the command
            if (pPtRes.Status == PromptStatus.Cancel) return;
            // Prompt for the end point
            pPtOpts.Message = "\nEnter the end point of the line: ";
            pPtOpts.UseBasePoint = true;
            pPtOpts.BasePoint = ptStart;
            pPtRes = documentoAtivo.Editor.GetPoint(pPtOpts);
            Point3d ptEnd = pPtRes.Value;
            if (pPtRes.Status == PromptStatus.Cancel) return;
            // Start a transaction
            using (Transaction acTrans = database.TransactionManager.StartTransaction())
            {
                BlockTable blockTable;
                BlockTableRecord blockTableRecord;
                // Open Model space for write
                blockTable = acTrans.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;
                blockTableRecord = acTrans.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                // Define the new line
                Line criarLinha = new Line(ptStart, ptEnd); criarLinha.SetDatabaseDefaults();
                // Add the line to the drawing
                blockTableRecord.AppendEntity(criarLinha); acTrans.AddNewlyCreatedDBObject(criarLinha, true);
                // Zoom to the extents or limits of the drawing
                documentoAtivo.SendStringToExecute("._zoom _all ", true, false, false);
                // Commit the changes and dispose of the transaction
                acTrans.Commit();
            }
        }
    }
}

