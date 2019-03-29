using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace TransmitLetter
{
  class WritterReader
  {
    public void Write(string Path,
                      string TemplatePath,
                      string SheetName,
                      int startRow,
                      int startColumn,
                      List<List<string>> Data,
                      string fileName,
                      string transmitNumber)
    {
      var excel = new Application();
      Worksheet ws;
      Workbooks wbs;
      Workbook wb;

      wbs = excel.Workbooks;
      wb = wbs.Open(TemplatePath);
      ws = wb.Sheets[SheetName];

      foreach (List<string> row in Data)
      {
        int _startColums = startColumn;
        foreach (string v in row)
        {
          ws.Cells[startRow, _startColums] = v;
          _startColums++;
        }
        startRow++;
      }

      if (!fileName.Contains("_CSV"))
      {
        ws.Cells[2, 3] = transmitNumber;
        ws.Cells[1, 9] = DateTime.Now.ToString("d");
      }

      string filePath = String.Format("{0}\\{1}", Path, fileName);

      if (File.Exists(filePath))
      {
        File.Delete(filePath);
      }
      wb.SaveAs(filePath);

      wbs.Close();
      excel.Quit();

      Marshal.ReleaseComObject(ws);
      Marshal.ReleaseComObject(wb);
      Marshal.ReleaseComObject(wbs);
      Marshal.ReleaseComObject(excel);
    }

    

    // Запись файлов CRS
    public void WriteCrs(string Path,
      List<List<string>> Data,
      string transmitNumber)
    {
      //путь к CRS
      GetPathsToTemplates ptt = new GetPathsToTemplates();
      string pathToCRS = ptt.getPathsToTemplates()[3];

      var excel = new Application();
      Worksheet ws;
      Workbooks wbs;
      Workbook wb;

      wbs = excel.Workbooks;
      wb = wbs.Open(pathToCRS);
      ws = wb.Sheets["Comment Review Sheet"];

      foreach (List<string> row in Data)
      {
        ws.Cells[7, 5] = transmitNumber;                //transmit no
        ws.Cells[7, 9] = DateTime.Now.ToString("d");    //date
        ws.Cells[22, 1] = row[1];                       //status
        ws.Cells[22, 2] = row[2];                       //doc No
        ws.Cells[22, 3] = row[3];                       //doc class
        ws.Cells[22, 4] = row[4];                       //doc title
        ws.Cells[22, 6] = row[5];                       //rev

        System.IO.Directory.CreateDirectory(String.Format("{0}\\CRS", Path));

        string filePath = String.Format("{0}\\CRS\\{1}", Path, row[0]);

        if (File.Exists(filePath))
        {
          File.Delete(filePath);
        }
        wb.SaveAs(filePath);
        
      }

      wbs.Close();
      excel.Quit();

      Marshal.ReleaseComObject(ws);
      Marshal.ReleaseComObject(wb);
      Marshal.ReleaseComObject(wbs);
      Marshal.ReleaseComObject(excel);
    }




    // Чтение таблицы Эксель: Путь, Лист
    public List<List<string>> Read(string Path, string SheetName)
    {
      var excel = new Application();
      Worksheet ws;
      Workbooks wbs;
      Workbook wb;

      wbs = excel.Workbooks;
      wb = wbs.Open(Path);
      ws = wb.Sheets[SheetName];

      Range usedRange = ws.UsedRange;
      object[,] values = usedRange.Value2;

      int rows = values.GetUpperBound(0);
      int cols = values.Length / rows;
     

      List<List<string>> Table = new List<List<string>>();

      for (int r = 1; r < rows; r++)
      {
        List<string> Row = new List<string>();

        for (int c = 1; c <= cols; c++)
        {
          Row.Add(Convert.ToString(values[r, c]));
        }
        Table.Add(Row);
      }

      wbs.Close();
      excel.Quit();

      Marshal.ReleaseComObject(usedRange);
      Marshal.ReleaseComObject(ws);
      Marshal.ReleaseComObject(wb);
      Marshal.ReleaseComObject(wbs);
      Marshal.ReleaseComObject(excel);

      return Table;
    }
  }
}
