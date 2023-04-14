using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Net;
using System.Web;
using System.Web.Http;

namespace DownloadFileCSV.Controllers
{
    public class Class1 : ApiController
    {
        /// <summary>
        /// 直接下載csv
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("GetExportCheckReportToCSV")]
        public HttpResponseMessage GetExportCheckReportToCSV()
        {
            // 建立一個 MemoryStream
            MemoryStream memoryStream = new MemoryStream();

            // 使用 StreamWriter 將文字寫入 MemoryStream
            StreamWriter streamWriter = new StreamWriter(memoryStream);
            streamWriter.WriteLine("Hello World!");
            streamWriter.Flush();

            // 將stream的位置設為0，表示從頭開始讀取或寫入資料。
            // 通常是在重新使用一個已經存在的MemoryStream物件時使用，以確保從頭開始讀取或寫入資料。
            memoryStream.Position = 0;



            string filename = DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
            // 建立一個 HttpResponseMessage
            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);

            // 將 MemoryStream 陣列寫入 HttpResponseMessage 的 Content 中
            // 設定 HttpResponseMessage 的 Content 屬性為一個 StreamContent 對象，並將要返回的文件流作為參數傳遞給 StreamContent 的構造函數。
            response.Content = new StreamContent(memoryStream);

            // 設定 Content-Type 屬性為要返回的文件的 MIME 類型。
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/csv");

            // 設定 Content-Disposition 屬性為 "attachment; filename=文件名"，其中 "文件名" 是要返回的文件的名稱。
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = filename
            };

            return response;
        }

    }
}