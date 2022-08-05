using System.Diagnostics;
using Excel1 = Microsoft.Office.Interop.Excel;


internal class Program
{
    static void Main(String[] args)
    {
        foreach(Process clsProcess in Process.GetProcesses())
        {
            if (clsProcess.ProcessName.Equals("Excel"))
            {
                clsProcess.Kill();
            }
        }

        Excel1.Application xlApp = new Excel1.Application();
        xlApp.Visible = true;
        xlApp.DisplayAlerts = false;

        Excel1.Workbook wb = xlApp.Workbooks.Open(@"F:\D\官司\证据0504\转账流水\excel格式\2.xlsx ");
        //这里要改
        Excel1.Worksheet ws = wb.Worksheets[1];

        Excel1.Range range1 = ws.Columns[8];
        range1.NumberFormat = "yyyy-MM-dd  hh:mm:ss";


        //这里要改
        for (int i = 3416; i > 8; i--)
        {   
            // Cells[行数,列号]
            if (ws.Cells[i,1].Value2 == null)
            {
  
                for (int j = 2;  j < 8; j++)
                {
                    if(j == 5)
                    {
                        continue;
                    }


                    if (ws.Cells[i - 1, j].Value2 == null)
                    {
                        ws.Cells[i - 1, j].Value2 = "";
                    }
                    if (ws.Cells[i , j].Value2 == null)
                    {
                        ws.Cells[i , j].Value2 = "";
                    }


                    ws.Cells[i - 1, j].Value2 = Convert.ToString(ws.Cells[i - 1, j].Value2) + " " + Convert.ToString(ws.Cells[i, j].Value2);

                    //ws.Cells[i - 1, j].Value2 += (" " + ws.Cells[i, j].Value2);
                }

                //if (ws.Cells[i,8].Value2 != null)
                //{
                 //   ws.Cells[i - 1, 8].Value2 = ws.Cells[i - 1, 8].Value2 + ws.Cells[i, 8].Value2;
                // ws.Cells[i - 1, 8].NumberFormat = "yyyy-MM-dd  hh:mm:ss";

                //}



                Excel1.Range range2 = ws.Rows[i];
                range2.Delete();
                //删除这行
            }
        }

        wb.SaveAs2(@"F:\D\官司\证据0504\转账流水\excel格式\2-1.xlsx");

        //这里要改
        wb.Close();
        xlApp.Workbooks.Close();
        xlApp.Quit();

        foreach (Process clsProcess in Process.GetProcesses())
        {
            if (clsProcess.ProcessName.Equals("Excel"))
            {
                clsProcess.Kill();
            }
        }


        Console.WriteLine("已经处理完毕");
        Console.ReadLine();




    }
}