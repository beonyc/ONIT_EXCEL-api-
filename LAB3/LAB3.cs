using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ONIT_3_4_5
{
   
 

    public class GetRequest
    {
        private HttpWebRequest _request;
        public string _address;

        public string Response { get; set; }
        public GetRequest(string address)
        {
            _address = address;
        }
        public void Run()
        {
            _request = (HttpWebRequest)WebRequest.Create(_address);
            _request.Method = "GET";
            try
            {
                HttpWebResponse response = (HttpWebResponse)_request.GetResponse();
                var stream = response.GetResponseStream();
                if (stream != null)
                {
                    Response = new StreamReader(stream).ReadToEnd();
                }

            }
            catch (Exception)
            {
            }

        }
    }
    internal class LAB3
    {
        public void ExportToExcel()
        {
            Excel.Application ex = new Excel.Application
            {
                Visible = true,
                SheetsInNewWorkbook = 1
            };

            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
            ex.DisplayAlerts = false;
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
            sheet.Name = $"Authot body for {DateTime.UtcNow.ToShortDateString()}";
            sheet.Columns.ColumnWidth = 20;
            sheet.Cells[1, 1] = "Автор";
            sheet.Cells[1, 2] = "Цитата";
          
            string url = "https://favqs.com/api/qotd";
          
            GetRequest request = new GetRequest(url);
            request.Run();
            var response = request.Response;

            var json = JObject.Parse(response);
            var quote = json["quote"];
            var Author = quote["author"];
            var Body = quote["body"];

            for (int i = 0; i < 1; i++)
            {
                sheet.Cells[i + 2, 1] = Author;
                sheet.Cells[i + 2, 2] = Body;
               

            }


        }

    }
}
