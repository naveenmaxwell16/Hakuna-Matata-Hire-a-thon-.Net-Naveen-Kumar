using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace Performance_Employee
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + ConfigurationManager.AppSettings["Json"];
            StreamReader r = new StreamReader(path);
            string jsonString = r.ReadToEnd();
            List<ExcelData> excelDatas = JsonConvert.DeserializeObject<List<ExcelData>>(jsonString);
            ReadExcel(excelDatas);
        }

        public static void ReadExcel(List<ExcelData> lstexcelDatas)
        {

            //Application excelApp = new Application();

            //if (excelApp == null)
            //{
            //    Console.WriteLine("Excel is not installed!!");
            //    return;
            //}

            //Workbook excelBook = excelApp.Workbooks.Open(@"D:\Hackathon Timesheet.xlsx");
            //_Worksheet excelSheet = (_Worksheet)excelBook.Sheets[1];
            //Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;
            //int rowCount = excelRange.Rows.Count;
            //int colCount = excelRange.Columns.Count;
            //List<ExcelData> lstexcelDatas = new List<ExcelData>();
            //Console.WriteLine("Excel reading..Please wait");
           
            //for (int i = 2; i <= rowCount; i++)
            //{
            //    ExcelData excelData = new ExcelData();
            //    //create new line
            //    Console.WriteLine(i);
            //    for (int j = 1; j <= colCount; j++)
            //    {
            //        //write the console
            //        if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
            //        {
            //            if (1 == j)
            //            {
            //                excelData.Date = excelRange.Cells[i, j].Value2.ToString();
            //            }
            //            else if (2 == j)
            //            {
            //                excelData.ProjectName = excelRange.Cells[i, j].Value2.ToString();
            //            }
            //            else if (3 == j)
            //            {
            //                excelData.Hours = excelRange.Cells[i, j].Value2.ToString();
            //            }
            //            else if (4 == j)
            //            {
            //                excelData.Owner = excelRange.Cells[i, j].Value2.ToString();
            //            }
            //            else if (5 == j)
            //            {
            //                excelData.Team = excelRange.Cells[i, j].Value2.ToString();
            //            }
            //            else
            //            {
            //                excelData.BillingStatus = excelRange.Cells[i, j].Value2.ToString();
            //            }
            //           // Console.Write(excelRange.Cells[i, j].Value2.ToString() + "\t");
            //        }
            //    }

            //    lstexcelDatas.Add(excelData);
               
            //}
            //var opt = new JsonSerializerOptions() { WriteIndented = true };
            //var strJson = System.Text.Json.JsonSerializer.Serialize<IList<ExcelData>>(lstexcelDatas, opt);

            FilterByTeam(lstexcelDatas);
           
            
            //excelApp.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            Console.ReadLine();
            //Console.WriteLine("Excel readed. successfully");
           
        }

        public static void FilterByTeam(List<ExcelData> lstexcelDatas)
        {
            //To filter by team to get the project.
            List<string>Team = lstexcelDatas.Select(x => x.Team).Distinct().ToList();
            //to filter by project to get the hours for project;
            for (int i = 0; i < Team.Count; i++)
            {
                string Teams = Team[i];
                Console.WriteLine();
                if (Teams != null)
                {
                    string teams = Team[i].ToString();
                    List<string> ProjectName = lstexcelDatas.Where(x => x.Team == teams).Select(x => x.ProjectName).Distinct().ToList();
                    //based on team and project to get a hours worked totally.
                    for (int j = 0; j < ProjectName.Count; j++)
                    {
                        string PrjctName = ProjectName[j].ToString();
                        List<string> Hours = lstexcelDatas.Where(x => x.Team == teams && x.ProjectName == PrjctName).Select(x => x.Hours).ToList();
                        //To calculate the hours based on respective project and team.
                        Console.WriteLine("Project Name =   " + PrjctName + "   " + "Teams =   " + teams);
                        Console.WriteLine("------------------------------------------------------------");
                        float HoursCal = 0;
                        for (int k = 0; k < Hours.Count; k++)
                        {
                            if (HoursCal == 0)
                            {
                                HoursCal = float.Parse(Hours[k]);
                            }
                            else
                            {
                                HoursCal = HoursCal + float.Parse(Hours[k]);
                            }
                        }
                        float mean = HoursCal / Hours.Count;
                        Console.WriteLine("Total Hours spend by "   + "Project Name =   " + PrjctName + "  and " + "Teams =   " + teams + "  is   " + mean);
                    }
                }
                
            }

            Console.WriteLine("5 Employees with the lowest efficiency");
            Console.WriteLine("---------------------------------------");
            List<ExcelData> ExcelDataFinal = new List<ExcelData>();
            ExcelDataFinal = lstexcelDatas.OrderByDescending(x => x.Hours).ToList();

            List<string> Owner = lstexcelDatas.Select(x => x.Owner).Distinct().ToList();
            float hrs = 0;
            List<EmpEff> empEffsLst = new List<EmpEff>();
            for (int z=0;z<Owner.Count;z++)
            {
                EmpEff empEff = new EmpEff();
                List<string> OwnerHours = lstexcelDatas.Where(x => x.Owner == Owner[z]).Select(x => x.Hours).ToList();
                for (int d = 0; d< OwnerHours.Count;d++)
                {
                    if (hrs == 0)
                    {
                        hrs = float.Parse(OwnerHours[z]);
                    }
                    else
                    {
                        hrs = hrs + float.Parse(OwnerHours[z]);
                    }
                }

                empEff.Hours = hrs;
                empEff.Owner = Owner[z];

                empEffsLst.Add(empEff);
            }
            //To print 5 employee efficient

            for (int q = 0;q< empEffsLst.Count;q++)
            {

            }
        }
    }
}
