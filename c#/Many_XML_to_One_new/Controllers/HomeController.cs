//using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Spire.Xls;
using System;

using OfficeOpenXml;
using Many_XML_to_One_new.Models;
using Spire.Xls.Collections;

namespace Many_XML_to_One_new.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Problem2_Solution()
        {
            //find the path for each file and create a list which contains each one of them
            string root = Server.MapPath("~");
            string parent = Path.GetDirectoryName(root);
            string grandParent = Path.GetDirectoryName(parent);

            var file1_position = grandParent + @"\input\sample1.xlsx";
            var file2_position = grandParent + @"\input\sample2.xlsx";
            var file3_position = grandParent + @"\input\sample3.xlsx";
            var file4_position = grandParent + @"\input\sample4.xlsx";
            string[] excelFiles = new String[] { file1_position, file2_position, file3_position, file4_position };

            //Set the final file location which is inside the output folder
            string _xlsSolutionFilePath = grandParent + @"\output\final-Problem2-solution.xlsx";

            //Create temporary WorkBooks which will tranfer our data to the final workbook
            Workbook sourcebook1 = new Workbook();
            Workbook sourcebook2 = new Workbook();

            //Create new xlsx file named final-Problem2-solution.xlsx and save it to output folder
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage ExcelPkg = new ExcelPackage();
            ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
            wsSheet1.Protection.IsProtected = false;
            //wsSheet1.Protection.AllowSelectLockedCells = false;
            ExcelPkg.SaveAs(new FileInfo(grandParent + @"\output\final-Problem2-solution.xlsx"));
            


            //Create our final workbook
            Workbook destinationBook = new Workbook();
            destinationBook.Version = ExcelVersion.Version2010;
            destinationBook.LoadFromFile(_xlsSolutionFilePath);
            destinationBook.Worksheets.Clear();

            //Create a list which contains tha id of each file and the names of its sheets (for example=> 1:PolicyData , 2:PolicyData , 3:SafetyData , 3:RandomSheet....)
            var sheetsNames = new List<SheetsNames>();

            //Save the names and ids to sheetsNames list ,loading in every loop a different sheet of each file
            for (int i = 0; i < excelFiles.Length; i++)
            {
                sourcebook2.LoadFromFile(excelFiles[i]);
                foreach (Worksheet sheet in sourcebook2.Worksheets)
                {
                    var sampleName = excelFiles[i].Split('\\').Last().Split('.').FirstOrDefault();
                    var sampleId = Convert.ToInt32(sampleName.Substring(sampleName.Length - 1));

                    sheetsNames.Add(new SheetsNames
                    {
                        sampleId = sampleId,
                        sheetName = sheet.Name
                    });
                    
                }
            }

            //Clear the temporary workbook i used ,so it doesn't have any data 
            sourcebook2.Worksheets.Clear();

            //Find the ids of the sheets than appear more than once and the ids of the sheets than appears only once in sheetsNames list
            var duplicates = sheetsNames.GroupBy(x => x.sheetName).Where(g => g.Count() > 1).Select(x => x.Key).ToList();
            var unique_sheets = sheetsNames.GroupBy(x => x.sheetName).Where(g => g.Count() == 1).Select(x => x.Key).ToList();

            //For each sheet with the same name, group together the data in a new sheet of the final workbook
            foreach (var dublicate in duplicates)
            {
                var dublicateBookId = sheetsNames.Where(x => x.sheetName == dublicate).Select(x => x.sampleId).ToList();
                //Sourcebook1 Loads the newest sheet of the dublicates
                sourcebook1.LoadFromFile(excelFiles[dublicateBookId[0] - 1]);

                for (var i = 0; i < dublicateBookId.Count; i++)
                {
                    Worksheet sourceSheet1 = (Worksheet)sourcebook1.Worksheets.Where(x => x.Name == dublicate).SingleOrDefault();

                    //Sourcebook2 Loads the oldest sheet of the dublicates
                    sourcebook2.LoadFromFile(excelFiles[dublicateBookId[i] - 1]);
                    Worksheet sourceSheet2 = (Worksheet)sourcebook2.Worksheets.Where(x => x.Name == dublicate).SingleOrDefault();

                    //the duplicates appears for the n time >1
                    if (i > 0 && duplicates.Contains(sourceSheet2.Name) )
                    {
                        //Add new sheet to the final workbook
                        destinationBook.Worksheets.Add(dublicate);
                        Worksheet destSheet = (Worksheet)destinationBook.Worksheets.Where(x => x.Name == dublicate).SingleOrDefault();
                        
                        //Define ranges
                        CellRange destRange = destSheet.Range[1,1,1,1];
                        CellRange sourceRange1 = sourceSheet1.Range[sourceSheet1.FirstRow , sourceSheet1.FirstColumn, sourceSheet1.LastRow, sourceSheet1.LastColumn];
                        CellRange sourceRange2 = sourceSheet2.Range[sourceSheet2.FirstRow + 1, sourceSheet2.FirstColumn, sourceSheet2.LastRow, sourceSheet2.LastColumn];

                        //Copy the ranges needed for each sourceSheet to the destinationBook and save
                        sourceSheet1.Copy(sourceRange1, destRange);
                        sourceSheet2.Copy(sourceRange2, destSheet.Range[1 + sourceSheet1.LastRow, sourceSheet1.FirstColumn, sourceSheet2.LastRow + sourceSheet1.LastRow, sourceSheet1.LastColumn]);
                        destinationBook.SaveToFile(_xlsSolutionFilePath, ExcelVersion.Version2010);

                        //Copy the column widths also from the source range to destination range foreach sourceRange but failed
                        //for (int m = 0; m < sourceRange1.Columns.Length; m++)
                        //{
                        //    destRange.Columns[m].ColumnWidth = sourceRange1.Columns[m].ColumnWidth;
                        //}

                        //for (int m = 0; m < sourceRange2.Columns.Length; m++)
                        //{
                        //    destRange.Columns[m].ColumnWidth = sourceRange2.Columns[m].ColumnWidth;
                        //}

                        //destinationBook.SaveToFile(_xlsSolutionFilePath, ExcelVersion.Version2010);

                    }


                }

            }

            //For each sheet with unique name save the data in a new sheet of the final workbook
            foreach (var unique_sheet in unique_sheets)
            {
                var unique_sheetId = sheetsNames.Where(x => x.sheetName == unique_sheet).Select(x => x.sampleId).ToList();
                sourcebook1.LoadFromFile(excelFiles[unique_sheetId[0] - 1]);

                Worksheet newSheet = (Worksheet)sourcebook1.Worksheets.Where(x => x.Name == unique_sheet).SingleOrDefault();

                //Copy file's sheet to our final workbook and save
                destinationBook.Worksheets.AddCopy(newSheet);
                destinationBook.SaveToFile(_xlsSolutionFilePath, ExcelVersion.Version2010);
            }

            try
            {
                //The final workbook Appears
                System.Diagnostics.Process.Start(_xlsSolutionFilePath);
            }
            catch
            {

            }

            //Remove all the sheets from the final workbook
            destinationBook.Worksheets.Clear();

            return RedirectToAction("Index");
        }
        
        public ActionResult Problem1_Solution()
        {
            string root = Server.MapPath("~");
            string parent = Path.GetDirectoryName(root);
            string grandParent = Path.GetDirectoryName(parent);

            var file1_position = grandParent + @"\input\sample1.xlsx";
            var file2_position = grandParent + @"\input\sample2.xlsx";
            var file3_position = grandParent + @"\input\sample3.xlsx";
            var file4_position = grandParent + @"\input\sample4.xlsx";

            //Create new xlsx file named final-Problem2-solution.xlsx and save it to output folder
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage ExcelPkg = new ExcelPackage();
            ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
            wsSheet1.Protection.IsProtected = false;
            //wsSheet1.Protection.AllowSelectLockedCells = false;
            ExcelPkg.SaveAs(new FileInfo(grandParent + @"\output\final-Problem1-solution.xlsx"));

            Workbook newbook = new Workbook();

            newbook.Version = ExcelVersion.Version2010;

            newbook.Worksheets.Clear();
            Workbook tempbook = new Workbook();

            string[] excelFiles = new String[] { file1_position, file2_position, file3_position , file4_position };

            string _xlsSolutionFilePath = grandParent + @"\output\final-Problem1-solution.xlsx";
            
            for (int i = 0; i<excelFiles.Length; i++)
            {
                tempbook.LoadFromFile(excelFiles[i]);

                foreach (Worksheet sheet in tempbook.Worksheets)
                {
                    sheet.Name = "sample" + (i + 1) + "-" + sheet.Name;
                    newbook.Worksheets.AddCopy(sheet);
                    
                }
            }

            newbook.SaveToFile(_xlsSolutionFilePath, ExcelVersion.Version2010);
            System.Diagnostics.Process.Start(_xlsSolutionFilePath);

            return RedirectToAction("Index");
            //Workbook wbToStream = new Workbook();
            //Worksheet sheetNew = wbToStream.Worksheets[0];
            ////wbToStream.save = "mpla.xlsx";
            ////sheetNew.Range["C10"].Text = "The sample demonstrates how to save an Excel workbook to stream.";
            //FileStream file_stream = new FileStream("final-Problem11.xlsx", FileMode.Create);
            //wbToStream.SaveToStream(file_stream);
            //file_stream.Close();
            //System.Diagnostics.Process.Start("final-Problem11.xlsx");


            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //workSheet = workBook.Worksheets[0];
            //workSheet.Range["A1"].Text = "This is a sample Excel dcouemnt       and created by Spire.XLS for .NET";
            //workBook.SaveToFile(_xlsFilePath + _xlsFileName);

            //Workbook wbToStream = new Workbook();

            //Worksheet sheet = wbFromStream.Worksheets[0];

            //sheet.Range["C10"].Text = "The sample demonstrates how to save an Excel workbook to stream.";

            //wbToStream.SaveToStream(file_stream);

            //////////////////////////////////////////////////////////////////////////////////////////////////////////

            //using (System.IO.FileStream fs = System.IO.File.Create(_xlsFilePath))
            //{
            //    for (byte i = 0; i < 100; i++)
            //    {
            //        fs.WriteByte(i);
            //    }
            //}

            //newbook.SaveToFile("Sample.xls", ExcelVersion.Version97to2003);

            //System.Diagnostics.Process.Start(newbook.FileName);

            //byte[] array = null;
            //using (var ms = new System.IO.MemoryStream())
            //{
            //    newbook.SaveToStream(ms, FileFormat.Version2010);
            //    ms.Seek(0, System.IO.SeekOrigin.Begin);
            //    array = ms.ToArray();
            //}

            //return File(array, "application / vnd.openxmlformats - officedocument.spreadsheetml.sheet", " Detail.xlsx");

            //package.Workbook.Properties.Title = "Attempts";
            //this.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //this.Response.AddHeader(
            //          "content-disposition",
            //          string.Format("attachment;  filename={0}", "ExcellData.xlsx"));
            //this.Response.BinaryWrite(package.GetAsByteArray());

        }


        //private void MergeXlsxFiles(string destXlsxFileName, params string[] sourceXlsxFileNames)
        //{
        //    Application excelApp = null;
        //    Workbook destWorkBook = null;
        //    var temppathForTarget = Path.Combine(Directory.GetCurrentDirectory(), Guid.NewGuid() + ".xls");

        //    if (File.Exists(temppathForTarget))
        //        File.Delete(temppathForTarget);

        //    try
        //    {
        //        excelApp = new Application
        //        {
        //            DisplayAlerts = false,
        //            SheetsInNewWorkbook = 3
        //        };
        //        destWorkBook = excelApp.Workbooks.Add();
        //        destWorkBook.SaveAs(temppathForTarget);


        //        foreach (var sourceXlsxFile in sourceXlsxFileNames)
        //        {
        //            var file = Path.Combine(Directory.GetCurrentDirectory(), sourceXlsxFile);
        //            var sourceWorkBook = excelApp.Workbooks.Open(file);

        //            foreach (Worksheet ws in sourceWorkBook.Worksheets)
        //            {
        //                var wSheet = destWorkBook.Worksheets[destWorkBook.Worksheets.Count];
        //                ws.Copy(wSheet);
        //                destWorkBook.Worksheets[destWorkBook.Worksheets.Count].Name =
        //                    ws.Name;
        //            }
        //            sourceWorkBook.Close(XlSaveAction.xlDoNotSaveChanges);
        //        }
        //        destWorkBook.Sheets[1].Delete();
        //        destWorkBook.SaveAs(destXlsxFileName);
        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //    finally
        //    {
        //        if (destWorkBook != null)
        //            destWorkBook.Close(XlSaveAction.xlSaveChanges);
        //        if (excelApp != null)
        //            excelApp.Quit();
        //    }
        //}
        public ActionResult Index()
        {
            //Convert_Many_ExcelFiles_To_One();
            //EEP();
            //ccc();
            //ManyWorkSheetsToOne();
            //Problem2_Solution();
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}