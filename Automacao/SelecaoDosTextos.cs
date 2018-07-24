//using Autodesk.AutoCAD.ApplicationServices;
//using Autodesk.AutoCAD.DatabaseServices;
//using Autodesk.AutoCAD.EditorInput;
//using Autodesk.AutoCAD.Geometry;
//using System.Collections.Generic;
//using TodosBlocos;
//using System;


//namespace Drenagem
//{
//    class SelecaoDosTextos
//    {        
//        private static List<TextoElevacao> _listaElevacao;       

//        public static void AreaSelecao(Point3d pMin, Point3d pMax, Point3d pCenter, double dFactor)
//        {
//            Document documentoAtivo = Application.DocumentManager.MdiActiveDocument;
//            Database database = documentoAtivo.Database;

//            int nCurVport = System.Convert.ToInt32(Application.GetSystemVariable("CVPORT"));
//            // Get the extents of the current space no points
//            // or only a center point is provided
//            // Check to see if Model space is current
//            if (database.TileMode == true)
//            {
//                if (pMin.Equals(new Point3d()) == true && pMax.Equals(new Point3d()) == true)
//                {
//                    pMin = database.Extmin;
//                    pMax = database.Extmax;
//                }
//            }
//            else
//            {
//                // Check to see if Paper space is current
//                if (nCurVport == 1)
//                {
//                    // Get the extents of Paper space
//                    if (pMin.Equals(new Point3d()) == true && pMax.Equals(new Point3d()) == true)
//                    {
//                        pMin = database.Pextmin;
//                        pMax = database.Pextmax;
//                    }
//                }
//                else
//                {
//                    // Get the extents of Model space
//                    if (pMin.Equals(new Point3d()) == true &&
//                    pMax.Equals(new Point3d()) == true)
//                    {
//                        pMin = database.Extmin;
//                        pMax = database.Extmax;
//                    }
//                }
//                // Start a transaction
//                using (Transaction acTrans = database.TransactionManager.StartTransaction())
//                {
//                    // Get the current view
//                    using (ViewTableRecord acView = documentoAtivo.Editor.GetCurrentView())
//                    {
//                        Extents3d eExtents;
//                        // Translate WCS coordinates to DCS
//                        Matrix3d matWCS2DCS;
//                        matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection);
//                        matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) *
//                        matWCS2DCS;
//                        matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist, acView.ViewDirection, acView.Target);
//                        // If a center point is specified, define the min and max
//                        // point of the extents
//                        // for Center and Scale modes
//                        if (pCenter.DistanceTo(Point3d.Origin) != 0)
//                        {
//                            pMin = new Point3d(pCenter.X - (acView.Width / 2), pCenter.Y - (acView.Height / 2), 0);
//                            pMax = new Point3d((acView.Width / 2) + pCenter.X,  (acView.Height / 2) + pCenter.Y, 0);
//                        }
//                        // Create an extents object using a line
//                        using (Line acLine = new Line(pMin, pMax))
//                        {
//                            eExtents = new Extents3d(acLine.Bounds.Value.MinPoint, acLine.Bounds.Value.MaxPoint);
//                        }
//                        // Calculate the ratio between the width and height of the current view
//                        double dViewRatio;
//                        dViewRatio = (acView.Width / acView.Height);
//                        // Tranform the extents of the view
//                        matWCS2DCS = matWCS2DCS.Inverse();
//                        eExtents.TransformBy(matWCS2DCS);
//                        double dWidth;
//                        double dHeight;
//                        Point2d pNewCentPt;
//                        // Check to see if a center point was provided (Center and Scale modes)
//                        if (pCenter.DistanceTo(Point3d.Origin) != 0)
//                        {
//                            dWidth = acView.Width;
//                            dHeight = acView.Height;
//                            if (dFactor == 0)
//                            {
//                                pCenter = pCenter.TransformBy(matWCS2DCS);
//                            }
//                            pNewCentPt = new Point2d(pCenter.X, pCenter.Y);
//                        }
//                        else // Working in Window, Extents and Limits mode
//                        {
//                            // Calculate the new width and height of the current view
//                            dWidth = eExtents.MaxPoint.X - eExtents.MinPoint.X;
//                            dHeight = eExtents.MaxPoint.Y - eExtents.MinPoint.Y;
//                            // Get the center of the view
//                            pNewCentPt = new Point2d(((eExtents.MaxPoint.X +
//                            eExtents.MinPoint.X) * 0.5),
//                            ((eExtents.MaxPoint.Y +
//                            eExtents.MinPoint.Y) * 0.5));
//                        }
//                        // Check to see if the new width fits in current window
//                        if (dWidth > (dHeight * dViewRatio)) dHeight = dWidth / dViewRatio;
//                        // Resize and scale the view
//                        if (dFactor != 0)
//                        {
//                            acView.Height = dHeight * dFactor;
//                            acView.Width = dWidth * dFactor;
//                        }
//                        // Set the center of the view
//                        acView.CenterPoint = pNewCentPt;// Set the current view
//                        documentoAtivo.Editor.SetCurrentView(acView);
//                    }
//                    // Commit the changes
//                    acTrans.Commit();
//                }
//            }
//            throw new NotImplementedException();
//        }
//        public void Selecao()
//        {           

//            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;

//            _listaElevacao = new List<TextoElevacao>();            

//            foreach (AtributosDoBloco atributo in lista)
//            {
//                TextoElevacao Elevacao1 = new TextoElevacao();

//                Point3d p1 = new Point3d(atributo.X - 7.5, atributo.Y + 7.5, 0);
//                Point3d p2 = new Point3d(atributo.X + 7.5, atributo.Y + 7.5, 0);
//                Point3d p3 = new Point3d(atributo.X + 7.5, atributo.Y - 7.5, 0);
//                Point3d p4 = new Point3d(atributo.X - 7.5, atributo.Y - 7.5, 0);

//                //Point3dCollection pntCol = new Point3dCollection();
//                //pntCol.Add(p1);
//                //pntCol.Add(p2);
//                //pntCol.Add(p3);
//                //pntCol.Add(p4);

//                AreaSelecao(p1, p3, new Point3d(), 1);

//                try
//                {
//                    PromptPointOptions ptOpts = new PromptPointOptions("\nSelecione as Elevações dos Tubos: ");
//                    PromptPointResult ptRes = editor.GetPoint(ptOpts);

//                    if (PromptStatus.OK == ptRes.Status)
//                    {
//                        int textoEncontrado = 0;

//                        Point3d p = ptRes.Value;
//                        Point3d ponto1 = new Point3d(p.X - 0.01, p.Y - 0.01, 0.0);
//                        Point3d ponto2 = new Point3d(p.X + 0.01, p.Y + 0.01, 0.0);

//                        TypedValue[] filterList = new TypedValue[2];
//                        filterList.SetValue(new TypedValue((int)DxfCode.Start, "MTEXT"), 0);
//                        filterList.SetValue(new TypedValue((int)DxfCode.Start, "TEXT"), 0);
//                        SelectionFilter selecaoFilter = new SelectionFilter(filterList);

//                        PromptSelectionResult res = editor.SelectCrossingWindow(ponto1, ponto2, selecaoFilter);
//                        SelectionSet ss = res.Value;
//                        if (ss != null)
//                        {
//                            foreach (ObjectId objId in res.Value.GetObjectIds())
//                            {
//                                string elevacaoTubulacao = textoEncontrado.ToString();
//                                string textoElevacaoInicial = "CA=";
//                                string textoElevacaoFinal = "CFC=";

//                                if (elevacaoTubulacao.Contains(textoElevacaoInicial))
//                                {
//                                    Elevacao1.ElevacaoInicial = elevacaoTubulacao;
//                                    _listaElevacao.Add(Elevacao1);
//                                }
//                                if (elevacaoTubulacao.Contains(textoElevacaoFinal))
//                                {
//                                    Elevacao1.ElevacaoFinal = elevacaoTubulacao;
//                                    _listaElevacao.Add(Elevacao1);
//                                }
//                                else
//                                    editor.WriteMessage("\nDid As Elevações não foram encontradas!");
//                            }

//                        }
//                        textoEncontrado++;
//                    }

//                }
//                catch (System.Exception e)

//                {

//                    editor.WriteMessage("\nException {0}.", e);

//                }
//            }
//        }

//    }

//}

