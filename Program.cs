using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenaricToExcel
{
    class Program
    {
        static void Main(string[] args)
        {

            PapulateExcel(getData());
        }


        private static List<CountriesWiseSummery> getData()
        {
            List<CountriesWiseSummery> list = new List<CountriesWiseSummery>();

            CountriesWiseSummery covid = new CountriesWiseSummery();
            covid.Country = "UAE";
            covid.NewConfirmed = 323;
            covid.TotalConfirmed = 22332233;
            covid.NewDeaths = 342;
            covid.TotalDeaths = 233323;
            covid.NewRecovered = 3232;
            covid.TotalRecovered = 322233;
            covid.Date = DateTime.Now.ToString();

            list.Add(covid);
            covid = new CountriesWiseSummery();
            covid.Country = "India";
            covid.NewConfirmed = 433;
            covid.TotalConfirmed = 54445;
            covid.NewDeaths = 666;
            covid.TotalDeaths = 433434;
            covid.NewRecovered = 544545;
            covid.TotalRecovered = 54454554;
            covid.Date = DateTime.Now.ToString();

            list.Add(covid);
            return list;
        }

        public static void PapulateExcel(List<CountriesWiseSummery> listData)
        {

            var excelApp = new Application();
            var workBook = excelApp.Workbooks.Add();
            var Sheet = (Worksheet)workBook.ActiveSheet;
           

            try
            {
               
                int counter = 1;

                foreach (var propInfo in listData[0].GetType().GetProperties())
                {
                    Sheet.Cells[1, counter] = propInfo.Name;
                    Sheet.Cells[1, counter].Font.Bold = true;
                    counter++;
                }

                int r = 0;

                foreach (var data in listData)
                {
                    int f = 1;
                    foreach (var propInfo in data.GetType().GetProperties())
                    {
                        try
                        {
                            Sheet.Cells[r + 2, f] = propInfo.GetValue(data);
                            f++;
                        }
                        catch (Exception d)
                        {

                        }
                    }
                    r++;
                }

                string path = "D:\\test.xlsx";

                workBook.SaveAs(path) ;
                workBook.Close(true);
                Console.WriteLine("Success Export");

            }
            catch (Exception ex)
            {
                workBook.Close(false);
                Console.WriteLine("Failed");
            }

            Console.ReadLine();
          
        }
       
    }
}
