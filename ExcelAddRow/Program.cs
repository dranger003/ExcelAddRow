using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelAddRow
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Count() != 1)
            {
                Console.WriteLine("Usage: {0} <path>", Environment.GetCommandLineArgs()[0]);
                return;
            }

            try
            {
                foreach (string fileName in Directory.GetFiles(args[0], "*.xlsx"))
                {
                    Console.WriteLine("Adding top row to worksheet \"Input\" for Excel file \"{0}\"...", fileName);

                    using (SpreadsheetDocument ssd = SpreadsheetDocument.Open(fileName, true))
                    {
                        WorkbookPart wbp = ssd.WorkbookPart;
                        Sheet sheet = wbp.Workbook.Descendants<Sheet>().Where(s => s.Name == "Input").FirstOrDefault();
                        WorksheetPart wsp = (WorksheetPart)ssd.WorkbookPart.GetPartById(sheet.Id.Value);

                        SheetData sd = wsp.Worksheet.GetFirstChild<SheetData>();
                        Row rr = sd.Descendants<Row>().Where(r => r.RowIndex == 2).FirstOrDefault();
                        Row nr = new Row() { RowIndex = 2 };

                        Cell a2 = new Cell() { CellReference = "A2", DataType = CellValues.InlineString };
                        {
                            InlineString ils = new InlineString();
                            ils.Append(new Text() { Text = "TypeGuessRows" });
                            a2.Append(ils);
                        }
                        Cell b2 = new Cell() { CellReference = "B2", DataType = CellValues.InlineString };
                        {
                            InlineString ils = new InlineString();
                            ils.Append(new Text() { Text = "0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789" });
                            b2.Append(ils);
                        }

                        nr.Append(a2);
                        nr.Append(b2);

                        CalculationChainPart ccp = wbp.CalculationChainPart;
                        if (ccp != null)
                        {
                            CalculationCell cc = ccp.CalculationChain.Descendants<CalculationCell>().Where(c => c.CellReference == "B2").FirstOrDefault();
                            if (cc != null)
                                cc.Remove();

                            if (ccp.CalculationChain.Count() == 0)
                                wbp.DeletePart(ccp);
                        }

                        foreach (Row rw in wsp.Worksheet.Descendants<Row>().Where(r => r.RowIndex.Value >= 2))
                        {
                            uint nri = Convert.ToUInt32(rw.RowIndex.Value + 1);
                            foreach (Cell cl in rw.Elements<Cell>())
                            {
                                string cr = cl.CellReference.Value;
                                cl.CellReference = new StringValue(cr.Replace(rw.RowIndex.Value.ToString(), nri.ToString()));
                            }
                            rw.RowIndex = new UInt32Value(nri);
                        }

                        sd.InsertBefore(nr, rr);

                        wsp.Worksheet.Save();
                    }
                }

                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //Excel.Application excel = new Excel.Application();
                //Excel.Workbook workbook = excel.Workbooks.Open(
                //    @"C:\DATA\NS_IMPORT\Network Standards - Managing Risk  - Canada.xlsx",
                //    0,
                //    false,
                //    5,
                //    String.Empty,
                //    String.Empty,
                //    true,
                //    Excel.XlPlatform.xlWindows,
                //    "\t",
                //    true,
                //    false,
                //    0,
                //    false,
                //    true,
                //    false);
                //Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item("Input");

                //worksheet.Rows.Insert(2, 1);
                //worksheet.Rows[2][0].Value = "TYPESELECTOR";
                //worksheet.Rows[2][1].Value = "0123456789";

                //excel.SaveWorkspace(@"C:\DATA\NS_IMPORT\OUTPUT.xlsx");
            }
            catch (Exception ex)
            {
                while (ex != null)
                {
                    Console.WriteLine(ex.Message);
                    ex = ex.InnerException;
                }
            }

            //Console.WriteLine("\nPress <any key> to continue.");
            //Console.ReadKey(true);
        }
    }
}
