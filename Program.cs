//Read Text Files
using System;
using System.IO;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;


namespace CheckDiskSpace
{
    class Class1
    {
        [STAThread]
        static void Main(string[] args)
        {
            String line;
            try
            {

                //string folderPath = @"Y:\ServersDiskSpace\ServerFiles"; //from local run
                string folderPath = @"\\ELCN-UTIL1\ServersDiskSpace\ServerFiles"; // from server run
                int month = DateTime.Now.Month;
                int day = DateTime.Now.Day;
                string YYYYMMDD;
                if (month < 10 && day < 10)
                {
                    YYYYMMDD = DateTime.Now.Year.ToString() + '0' + DateTime.Now.Month.ToString() + '0' + DateTime.Now.Day.ToString();
                }
                else if (month < 10 && day > 10)
                {
                    YYYYMMDD = DateTime.Now.Year.ToString() + '0' + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString();
                }
                else if (month > 10 && day < 10)
                {
                    YYYYMMDD = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + '0' + DateTime.Now.Day.ToString();
                }
                else
                {
                    YYYYMMDD = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString();
                }
                //string write_path = @"Y:\\ServersDiskSpace\Logs\serverdiskspace_" + YYYYMMDD + ".xlsx"; // from local run
                //string write_path = @"Y:\\ServersDiskSpace\Logs\serverdiskspace_" + YYYYMMDD + ".csv"; // from local csv run
                string write_path = @"\\ELCN-UTIL1\ServersDiskSpace\Logs\serverdiskspace_" + YYYYMMDD + ".xlsx"; // from server run
                #region EPPLus Info
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var epplusexcelfile = new FileInfo(write_path);
                using var package = new ExcelPackage(epplusexcelfile);
                var ws = package.Workbook.Worksheets.Add("ServersFreeDiskSpace_wb");
                List<ServerInfo> list = new List<ServerInfo>();


                #endregion EPPlus Info End

                DirectoryInfo directoryInfo = new DirectoryInfo(folderPath);
                FileInfo[] files = directoryInfo.GetFiles("*.txt");

                #region start reading files 
                foreach (FileInfo file in files)
                {

                    StreamReader sr = new StreamReader(file.FullName);
                    StreamWriter sw;
                    string fileName = file.FullName.Substring(folderPath.Length);
                    #region write to text(csv) file
                    //if (File.Exists(write_path))
                    //{

                    //    sw = new StreamWriter(write_path, true);
                    //    sw.Write(Path.GetFileNameWithoutExtension(fileName) + ",");
                    //    sw.Write(DateTime.Now + ",");

                    //}
                    //else
                    //{
                    //    sw = new StreamWriter(write_path);
                    //    sw.Write(Path.GetFileNameWithoutExtension(fileName) + ",");
                    //    sw.Write(DateTime.Now + ",");

                    //}
                    #endregion end write to text file

                    ServerInfo server = new ServerInfo();
                    server.ServerName = Path.GetFileNameWithoutExtension(fileName);
                    server.RunTime = DateTime.Now;

                    //Read the first line of text
                    line = sr.ReadLine();

                    //Continue to read until you reach end of file
                    while (line != "EOF")
                    {

                        line = sr.ReadLine();
                        Console.WriteLine(line);
                        string[] items = line.Split('\t', '\n');

                        #region write to a file               
                        foreach (string item in items)
                        {

                            if (item.Contains("FileSystem"))
                            {

                                int index_FS = item.IndexOf("FileSystem");
                                string temp1 = item.Substring(0, index_FS - 1);
                                string current_drive = temp1.Substring(0, 1);
                                int index_Space = temp1.Trim().LastIndexOf("");
                                if (index_Space > 1)
                                {
                                    string freeSpace = temp1.Substring(index_Space - 5);
                                    //sw.Write(current_drive.Trim() + "," + freeSpace.TrimEnd() + ",");                                    
                                    DiskInfo currentDiskInfo = new DiskInfo();
                                    currentDiskInfo.DiskName = current_drive.Trim();
                                    currentDiskInfo.FreeSpace = freeSpace.TrimEnd();
                                    server.DiskInfo.Add(currentDiskInfo);
                                }

                            }

                        }
                        #endregion end of writing a file
                    }
                    //sw.WriteLine();
                    sr.Close();
                    list.Add(server);
                }
                #region write to memory list 
                foreach (ServerInfo item in list)
                {
                    Console.Write(item.ServerName + " " + item.RunTime + " ");
                    foreach (DiskInfo di in item.DiskInfo)
                    {
                        Console.Write(di.DiskName + " " + di.FreeSpace);

                    }
                    Console.WriteLine();
                }
                #endregion end write to memory list

                SaveExcelFile(list, epplusexcelfile);
                //  package.Save();
                #endregion end of reading files 
            } // end try
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);

            }
            finally
            {
                Console.WriteLine("Executing finally block.");
            }


        }
        #region Save To Excel File 
        public static void SaveExcelFile(List<ServerInfo> serverInfo, FileInfo file)
        {

            using var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("ServersFreeDiskSpace_wb");

            ws.Cells["A1"].Value = "Servers Free Disk Space Report";
            ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Row(1).Style.Font.Size = 24;
            ws.Row(1).Style.Font.Color.SetColor(Color.DarkBlue);
            ws.Cells["A1:R1"].Merge = true;

            ws.Cells[2, 1].Value = "SERVER";
            ws.Cells[2, 1].Style.Font.Bold = true;
            ws.Cells[2, 2].Value = "TIME RUN";
            ws.Cells[2, 2].Style.Font.Bold = true;
            ws.Cells["C2:R2"].Merge = true;
            ws.Cells[2, 3].Value = "AVAILABLE DISK SPACE ( GB )";
            ws.Cells[2, 3].Style.Font.Bold = true;
            ws.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#939c80");
            ws.Cells[2, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(colFromHex);
            ws.Cells[2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[2, 2].Style.Fill.BackgroundColor.SetColor(colFromHex);

            ws.Cells[2, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[2, 3].Style.Fill.BackgroundColor.SetColor(colFromHex);
            int row = 3;
            int col = 1;
            foreach (ServerInfo server in serverInfo)
            {
                ws.Cells[row, col++].Value = server.ServerName;
                ws.Cells[row, col++].Value = server.RunTime.ToShortDateString() + ' ' + server.RunTime.ToShortTimeString();
                int len = server.DiskInfo.Count;

                foreach (DiskInfo diskInfo in server.DiskInfo)
                {
                    //switch (Convert.ToDouble(diskInfo.FreeSpace))
                    //{
                    //    case < 20.0 :
                    //        ws.Cells[row, col++].Value = diskInfo.DiskName;
                    //        ColorCells(ws, row, col);
                    //        ws.Cells[row, col++].Value = diskInfo.FreeSpace; break;
                    //    default:
                    //        ws.Cells[row, col++].Value = diskInfo.DiskName;
                    //        ws.Cells[row, col++].Value = diskInfo.FreeSpace; break;
                    //}
                    if (diskInfo.DiskName.Equals("C") && Convert.ToDouble(diskInfo.FreeSpace) < 20.0)
                    {
                            ws.Cells[row, col++].Value = diskInfo.DiskName;                            
                            ColorCells(ws, row, col);
                            ws.Cells[row, col++].Value = diskInfo.FreeSpace;
                    }
                    else
                    { 
                            ws.Cells[row, col++].Value = diskInfo.DiskName;                           
                            ws.Cells[row, col++].Value = diskInfo.FreeSpace;
                    }
                    
                }
                row++; // just increase the row to print each server's information
                col = 1; // reset the column back to 1
            }          
            package.Save();
        }
        #endregion End Save To Excel File
        #region Coolor Cells
        public static void ColorCells(ExcelWorksheet ws, int row, int col)
        {
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFF00");
            ws.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row, col].Style.Fill.BackgroundColor.SetColor(colFromHex);
            ws.Cells[row, col].Style.Font.Color.SetColor(Color.Red);
        }
        #endregion end Color Cells
    }    
}
