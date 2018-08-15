using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Threading;

public class SynchronousSocketClient
{
    static Dictionary<int, double> lstData;
    public static void StartClient()
    {
        byte[] bytes = new byte[1024];
        
        try
        {
            IPHostEntry ipHostInfo = Dns.GetHostEntry(Dns.GetHostName());
            IPAddress ipAddress = ipHostInfo.AddressList[0];
            IPEndPoint remoteEP = new IPEndPoint(ipAddress, 11000);
            
            Socket sender = new Socket(ipAddress.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
            lstData = new Dictionary<int, double>();
            try
            {
                sender.Connect(remoteEP);

                Console.WriteLine("Socket connected to {0}",
                    sender.RemoteEndPoint.ToString());
                for (int i = 1; i <= 360; i++)
                {
                    byte[] msg = Encoding.ASCII.GetBytes(i.ToString());
                    int bytesSent = sender.Send(msg);
                    int bytesRec = sender.Receive(bytes);
                    string output = Encoding.ASCII.GetString(bytes, 0, bytesRec);
                    Console.WriteLine("{0}", output);


                    lstData.Add(i, Double.Parse(output));
                }
                Thread thread = new Thread(FormExcelChart);
                thread.Start();

                sender.Shutdown(SocketShutdown.Both);
                sender.Close();
            }
            catch (ArgumentNullException ane)
            {
                Console.WriteLine("ArgumentNullException : {0}", ane.ToString());
            }
            catch (SocketException se)
            {
                Console.WriteLine("SocketException : {0}", se.ToString());
            }
            catch (Exception e)
            {
                Console.WriteLine("Unexpected exception : {0}", e.ToString());
            }

        }
        catch (Exception e)
        {
            Console.WriteLine(e.ToString());
        }
    }
    static void FormExcelChart()
    {
        Application excel;
        Workbook worKbooK;
        Worksheet worKsheeT;
        try
        {
            excel = new Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            worKbooK = excel.Workbooks.Add(Type.Missing);

            worKsheeT = (Worksheet)worKbooK.ActiveSheet;
            int i = 1;
            foreach (var str in lstData)
            {
                worKsheeT.Cells[i, 1] = str.Key;
                worKsheeT.Cells[i++, 2] = str.Value;
            }

            Range chartRange;

            ChartObjects xlCharts = (ChartObjects)
               worKsheeT.ChartObjects(Type.Missing);
            ChartObject myChart = (ChartObject)
               xlCharts.Add(110, 15, 468, 315);
            Chart chartPage = myChart.Chart;

            chartRange = worKsheeT.get_Range("B1", "B360");
            chartPage.SetSourceData(chartRange, Type.Missing);
            chartPage.ChartType = XlChartType.xlLine;

            Series ser = (Series)chartPage.SeriesCollection(1);

            ser.Values = worKsheeT.Range[worKsheeT.Cells[1, 2], worKsheeT.Cells[360, 2]];
            ser.XValues = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[360, 1]];
            chartPage.HasLegend = false;

            Axis vertAxis = (Axis)chartPage.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
            vertAxis.HasMajorGridlines = true;
            vertAxis.MaximumScaleIsAuto = false;
            vertAxis.MaximumScale = 80;
            vertAxis.MinimumScaleIsAuto = false;
            vertAxis.MinimumScale = -80;
            vertAxis.MajorUnit = 20;
            vertAxis.MinorUnit = 4;
            chartPage.Export(@"E:\Coding\CoderBrain\ClientConsole\ClientConsole\TanChart.bmp",
               "BMP", Type.Missing);

            worKbooK.SaveAs(@"E:\Coding\CoderBrain\ClientConsole\ClientConsole\TanChart.xlsx");
            excel.Quit();
        }
        catch (Exception e)
        {
            Console.WriteLine(e.ToString());
        }
    }
}
