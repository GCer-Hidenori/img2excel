using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace img2excel
{
    class Excel
    {
        Application @excelApp;
        //Workbook @book;
        Worksheet @sheet;
        public Excel()
        {
            try
            {
                @excelApp = (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                @excelApp = new Application();
                @excelApp.Application.Visible = true;
            }
            Workbook book = @excelApp.Workbooks.Add();
            @sheet = (Worksheet)book.Worksheets.Item[1];
            @sheet.Cells.Select();
            @excelApp.Selection.ColumnWidth = 1;
            @excelApp.Selection.RowHeight = 10;
        }
        public void draw(int x,int y,Color color)
        {
            try{
                @excelApp.EnableEvents = false;
                Range range = (Range)@sheet.Cells[y + 1, x + 1];
                range.Interior.Color = color;
	        }catch(Exception e){
                Console.Error.WriteLine("Error while drawing x={0} y={1}.\n{2}\n{3}", x, y);
                throw e;
            }finally{
                @excelApp.EnableEvents = true;
            }
           
        }
    }
}
