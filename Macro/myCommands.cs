// (C) Copyright 2022 by  
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Macro
{
    public class MyCommands
    {
        [CommandMethod("Layers_Turns_Off_True")]
        public static void Layers_Turns_Off_True()
        {
            LayersForm lf = new LayersForm();
            lf.Show();
        }
        public void LAYERS(string filepath, Document doc)
        {
            Database db = doc.Database;
            List<string> layers = new List<string>();
            Excel.Application oExcel = new Excel.Application();
            Excel.Workbook WB = oExcel.Workbooks.Open(filepath);
            string ExcelWorkbookname = WB.Name;
            int worksheetcount = WB.Worksheets.Count;
            Excel.Worksheet wks = WB.Worksheets[1];
            string firstworksheetname = wks.Name;
            Excel.Range xlRange = wks.UsedRange;
            foreach (Excel.Range item in xlRange.Rows.Cells)
            {
                var address = item.Address;
                var value = item.Value;
                layers.Add(value);
            }
            var layercount = layers.Count;
            foreach (var layer in layers)
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    doc.LockDocument();
                    LayerTable lyTab = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    foreach (ObjectId lyID in lyTab)
                    {
                        LayerTableRecord lytr = trans.GetObject(lyID, OpenMode.ForRead) as LayerTableRecord;
                        if (lytr.Name == layer)
                        {
                            lytr.UpgradeOpen();
                            lytr.IsOff = true;
                            trans.Commit();
                            doc.Editor.WriteMessage("\nLayer " + lytr.Name + " has been turned Off.");
                            break;
                        }
                        else
                        {
                            doc.Editor.WriteMessage("\nLayer not found.");
                        }
                    }
                }
            }
            oExcel.Workbooks.Close();
        }

        [CommandMethod("Canning_Blocks_True")]
        public void Canning_Blocks_True()
        {
            MainForm mf = new MainForm();
            mf.Show();
        }
        public void CAN(string num, string discipline, PromptPointResult ppr, string pathFile)
        {
            string miniDiscipline;
            switch (discipline)
            {
                case "SANITARY":
                    miniDiscipline = "San_";
                    break;
                case "STORM":
                    miniDiscipline = "Storm_";
                    break;
                case "DCW":
                    miniDiscipline = "DCW_";
                    break;
                case "GAS":
                    miniDiscipline = "Gas_";
                    break;
                case "VENT":
                    miniDiscipline = "Vent_";
                    break;
                case "HPS":
                    miniDiscipline = "HPS_";
                    break;
                case "HPR":
                    miniDiscipline = "HPR_";
                    break;
                case "HWS":
                    miniDiscipline = "HWS_";
                    break;
                case "HWR":
                    miniDiscipline = "HWR_";
                    break;
                case "CWS":
                    miniDiscipline = "CWS_";
                    break;
                case "CWR":
                    miniDiscipline = "CWR_";
                    break;
                case "DHW":
                    miniDiscipline = "DHW_";
                    break;
                case "DHWR":
                    miniDiscipline = "DHWR_";
                    break;
                case "COND":
                    miniDiscipline = "COND_";
                    break;
                default:
                    miniDiscipline = "";
                    break;
            }
            List<string> CadFile = new List<string>()
            {
                miniDiscipline + num + ".dwg"
            };
            CopyAndPasteExternalFile(CadFile, pathFile, ppr);
        }
        public static void CopyAndPasteExternalFile(List<string> CadFamily, string path, PromptPointResult ppr)
        {
            FileInfo[] Files = new DirectoryInfo(path).GetFiles("*.dwg");
            foreach (FileInfo file in Files)
            {
                foreach (var cadFamily in CadFamily)
                {
                    if (file.Name == cadFamily)
                    {
                        var fileName = Path.GetFileName(file.FullName);
                        string dwgFlpath = path + fileName;
                        Document docCurrent = Application.DocumentManager.MdiActiveDocument;
                        Database dbCurrent = docCurrent.Database;
                        Editor ed = docCurrent.Editor;
                        using (Database dbSource = new Database(false, true))
                        {
                            dbSource.ReadDwgFile(dwgFlpath, FileOpenMode.OpenForReadAndAllShare, false, null);
                            IdMapping mapping = new IdMapping();
                            using (Transaction tr = dbSource.TransactionManager.StartTransaction())
                            {
                                docCurrent.LockDocument();
                                ObjectId sourceMsId = SymbolUtilityServices.GetBlockModelSpaceId(dbSource);
                                ObjectId destDbMsId = SymbolUtilityServices.GetBlockModelSpaceId(dbCurrent);
                                ObjectIdCollection sourceIds = new ObjectIdCollection();
                                BlockTable bt = tr.GetObject(dbSource.BlockTableId, OpenMode.ForRead) as BlockTable;
                                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                                foreach (ObjectId id in btr)
                                {
                                    sourceIds.Add(id);
                                }
                                dbSource.WblockCloneObjects(sourceIds, destDbMsId, mapping, DuplicateRecordCloning.Replace, false);
                                tr.Commit();
                            }
                            using (Transaction tr2 = dbCurrent.TransactionManager.StartTransaction())
                            {
                                docCurrent.LockDocument();
                                Point3d pt1 = ppr.Value;
                                BlockTable btCurrent = tr2.GetObject(dbCurrent.BlockTableId, OpenMode.ForRead) as BlockTable;
                                BlockTableRecord btrCurrent = tr2.GetObject(btCurrent[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                                BlockTableRecordEnumerator iter = btrCurrent.GetEnumerator();
                                ObjectId lastObjId = new ObjectId();
                                while (iter.MoveNext())
                                {
                                    lastObjId = iter.Current;
                                }
                                BlockReference blk = tr2.GetObject(lastObjId, OpenMode.ForWrite) as BlockReference;
                                blk.Position = pt1;
                                tr2.Commit();
                            }
                        }
                    }
                }
            }
        }
    }

}
