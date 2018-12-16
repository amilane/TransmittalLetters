using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace TransmitLetter
{
  class WritterReader
  {
    public void Write(string Path)
    {
      

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

      int rows = values.GetUpperBound(0) + 1;
      int cols = values.Length / rows;

      List<List<string>> Table = new List<List<string>>();

      for(int r = 1; r < rows; r++)
      {
        List<string> Row = new List<string>();

        for (int c = 1; c < cols; c++)
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
