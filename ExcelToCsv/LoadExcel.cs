using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

class CLoadExcel
{
    Excel.Application Application = new Excel.Application();

    public void LoadExcel(System.IO.FileInfo FileInfo)
    {
        string DirectoryStr = FileInfo.Name.Split('.')[0];
        Directory.CreateDirectory(DirectoryStr);

        Excel.Workbook WorkBook = Application.Workbooks.Open(FileInfo.FullName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

        for (int i = 0; i < WorkBook.Worksheets.Count; ++i)
        {
            Excel.Worksheet WorkSheet = WorkBook.Worksheets.get_Item(i + 1);
            if (WorkSheet.Name.Contains("#"))
                continue;

            FileStream fileStream = new FileStream(DirectoryStr + "/" + WorkSheet.Name + ".csv", FileMode.OpenOrCreate);
            //StreamWriter Writer = new StreamWriter(fileStream);
            TextWriter Writer = new StreamWriter(fileStream, System.Text.Encoding.Unicode);

            Excel.Range Range = WorkSheet.UsedRange;

            string CsvStr = string.Empty;
            for(int RowIdx=1; RowIdx <= Range.Rows.Count; ++RowIdx)
            {
                for (int ColumnIdx = 1; ColumnIdx <= Range.Columns.Count; ++ColumnIdx)
                {
                    if ((Range.Cells[RowIdx, ColumnIdx] as Excel.Range).Value2 != null)
                    { 
                        var Value = (Range.Cells[RowIdx, ColumnIdx] as Excel.Range).Value2;
                        CsvStr += Value;
                        if (ColumnIdx != Range.Columns.Count)
                        {
                            CsvStr += ",";
                        }
                    }
                    
                }
                CsvStr += System.Environment.NewLine;
            }
            Writer.Write(CsvStr);
            Writer.Close();
            fileStream.Close();
        }

        WorkBook.Close();
        
        
    }
}
