using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Microsoft.Office.Interop.Excel;

public class SynchronousSocketClient
{
    public static void StartClient()
    {
        // Data buffer for incoming data.  
        byte[] bytes = new byte[1024];

        // Connect to a remote device.  
        try
        {
            // Establish the remote endpoint for the socket.  
            // This example uses port 11000 on the local computer.  
            IPHostEntry ipHostInfo = Dns.GetHostEntry(Dns.GetHostName());
            IPAddress ipAddress = ipHostInfo.AddressList[0];
            IPEndPoint remoteEP = new IPEndPoint(ipAddress, 11000);

            // Create a TCP/IP  socket.  
            Socket sender = new Socket(ipAddress.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
            Dictionary<int, double> lstData = new Dictionary<int, double>();
            string message = "HKB";
            // Connect the socket to the remote endpoint. Catch any errors.  
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
                //while (message != "close")
                //{
                //    // Encode the data string into a byte array.  
                //    byte[] msg = Encoding.ASCII.GetBytes("Client : "+message);

                //    // Send the data through the socket.  
                //    int bytesSent = sender.Send(msg);

                //    // Receive the response from the remote device.  
                //    int bytesRec = sender.Receive(bytes);
                //    Console.WriteLine("Server : {0}",
                //    Encoding.ASCII.GetString(bytes, 0, bytesRec));
                //    message = Console.ReadLine();
                //}

                sender.Shutdown(SocketShutdown.Both);
                sender.Close();

                //excel operations
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
                    //worKsheeT.Name = "StudentRepoertCard";
                    int i = 1;
                    foreach(var str in lstData)
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
                    
                    ser.Values = worKsheeT.Range[worKsheeT.Cells[1,2], worKsheeT.Cells[360, 2]];
                    ser.XValues = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[360, 1]];
                    //ser.Values = new int[] { 80, 60, 40, 20, 0, -20, -40, -60, -80 };
                    chartPage.HasLegend = false;

                    Axis vertAxis = (Axis)chartPage.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                    vertAxis.HasMajorGridlines = true; 
                    vertAxis.MaximumScaleIsAuto = false;
                    vertAxis.MaximumScale = 80; 
                    vertAxis.MinimumScaleIsAuto = false;
                    vertAxis.MinimumScale = -80;
                    vertAxis.MajorUnit = 20;
                    vertAxis.MinorUnit = 4;

                    // Export chart as picture file
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
}
