namespace TestWebBisness {

    #region using

    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Text;
    using System.Xml;
    using OfficeOpenXml;
    using OfficeOpenXml.Style;
    using System.Drawing;
    #endregion

    /// <summary>
    /// Http 请求
    /// </summary>
    public class HttpRequest {

        public HttpWebResponse PostReqeust(string url, string postDataStr) {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.ContentLength = Encoding.UTF8.GetByteCount(postDataStr);
            Stream myRequestStream = request.GetRequestStream();
            StreamWriter myStreamWriter = new StreamWriter(myRequestStream, Encoding.GetEncoding("gb2312"));
            myStreamWriter.Write(postDataStr);
            myStreamWriter.Close();

            return (HttpWebResponse)request.GetResponse();
        }

        /// <summary>
        /// post请求
        /// </summary>
        /// <param name="url"></param>
        /// <param name="postDataStr"></param>
        /// <returns>返回请求结果</returns>
        public static string GetPostReqeustString(string url, string postDataStr) {
            try {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Timeout = 1000;//超时
                request.Method = "POST";
                request.ContentType = "application/json";
                request.ContentLength = Encoding.Default.GetByteCount(postDataStr);
                Stream myRequestStream = request.GetRequestStream();
                StreamWriter myStreamWriter = new StreamWriter(myRequestStream, Encoding.Default);
                myStreamWriter.Write(postDataStr);
                myStreamWriter.Close();

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("fetch->{0}", url);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream myResponseStream = response.GetResponseStream();
                //根据服务器返回修改编码格式
                StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("gb2312"));
                return myStreamReader.ReadToEnd();
            } catch(Exception ex) {
                return string.Format("Code:\"{0}\"", -1);
            }
        }

        public void TestFriendsService() {

            string baseUrl = "http://115.231.94.9:8001/Service/FriendsService.asmx/{0}";
            string url = string.Format(baseUrl, "FriendList");
            string postData = "{ userId:7,SessionKey:\"-1941260349\",childId:11 }";
            string result = GetPostReqeustString(url, postData);


        }

    }

    /// <summary>
    /// Http Post class
    /// </summary>
    public class HttpPost {

        public string Url { get; set; }

        public string Method { get; set; }

        public Dictionary<string, string> param { get; set; }

        public string Result { get; set; }

        public HttpPost() {
            if(param == null) {
                param = new Dictionary<string, string>();
            }
        }

        public HttpPost(string url, string method, Dictionary<string, string> param) {
            this.Url = url;
            this.Method = method;
            this.param = param;
        }

        /// <summary>
        /// 组合Web service URL
        /// </summary>
        /// <returns></returns>
        public string ToFullUrl() {
            return string.Format("{0}/{1}", this.Url, this.Method);
        }

        /// <summary>
        /// 生成post数据格式
        /// </summary>
        /// <returns></returns>
        public string GetPostData() {
            var entries = this.param.Select(d => string.Format("{0}:\"{1}\"", d.Key, string.Join(",", d.Value)));
            return "{" + string.Join(",", entries) + "}";
        }

        public string GetName() {
            int i = this.Url.LastIndexOf('/') + 1;
            return string.Format("{0}_{1}", this.Url.Substring(i, this.Url.Length - i), this.Method);
        }
    }


    class Program {
        static void Main(string[] args) {
            //Bisness bisness = new Bisness();
            //bisness.TestFriendsService();

            #region 解析

            List<HttpPost> https = new List<HttpPost>();
            string xmlPath = "WebServices.xml";
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlPath);
            XmlNode root = doc.SelectSingleNode("webService");
            XmlNodeList serviceNode = root.SelectNodes("service");
            foreach(XmlNode service in serviceNode) {
                //读取URL
                string url = service.Attributes["url"].Value;
                XmlNodeList methodNode = service.SelectNodes("method");
                foreach(XmlNode method in methodNode) {
                    //读取方法名称
                    string methodName = method.Attributes["name"].Value;
                    XmlNode param = method.SelectSingleNode("param");
                    Dictionary<string, string> paramDic = new Dictionary<string, string>();
                    //处理参数
                    foreach(XmlNode item in param.ChildNodes) {
                        string nodeName = item.Name;
                        string nodeType = item.Attributes["type"].Value;
                        string nodeValue = item.InnerText;
                        //处理特殊参数
                        if(nodeName == "userId") {
                            nodeValue = "353";
                        } else if(nodeName == "SessionKey") {
                            nodeValue = "-884070913";
                        }

                        //打印测试
                        //Console.WriteLine("Url:{4},方法名称:{3},参数:{0}, 类型:{1}, 值:{2}", nodeName, nodeType, nodeValue, methodName, url);
                        paramDic.Add(nodeName, nodeValue);
                    }

                    https.Add(new HttpPost(url, methodName, paramDic));
                }
            }

            #endregion

            #region Http请求

            foreach(HttpPost http in https) {
                string url = http.ToFullUrl();
                string postData = http.GetPostData();
                string result = HttpRequest.GetPostReqeustString(url, postData);
                //"Code":"0" 
                int codeIndex = result.IndexOf("Code") + 7;
                string getResult = result.Substring(codeIndex, 1);
                // getResult==0 服务器返回成功
                http.Result = getResult;
            }

            #endregion

            #region 报告

            GenerateReport(https);
            #endregion

            Console.WriteLine("按任意键结束..");
            Console.ReadLine();
        }

        static void GenerateReport(List<HttpPost> https) {
            var stream = new MemoryStream();
            using(var xlpackage = new ExcelPackage(stream)) {
                foreach(HttpPost http in https) {
                    var worksheet = xlpackage.Workbook.Worksheets.Add(http.GetName());
                    worksheet.Cells["A1"].Value = "Url：";
                    worksheet.Cells["B1"].Value = http.Url;
                    worksheet.Cells["A2"].Value = "方法：";
                    worksheet.Cells["B2"].Value = http.Method;
                    worksheet.Cells["A3"].Value = "参数:";
                    worksheet.Cells["D3"].Value = "测试结果：";
                    worksheet.Cells["E3"].Value = http.Result == "0" ? "成功" : "失败";
                    worksheet.Cells["A4"].Value = "名称:";
                    worksheet.Cells["B4"].Value = "类型:";
                    worksheet.Cells["C4"].Value = "值:";
                    worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    worksheet.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Column(2).AutoFit();
                    int row = 5;
                    foreach(var p in http.param) {
                        int col = 1;
                        worksheet.Cells[row, col].Value = p.Key;
                        col += 1;
                        worksheet.Cells[row, col].Value = p.Value;

                        row += 1;
                    }

                    //worksheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 32, 96));
                }

                xlpackage.Save();
            }

            byte[] bytes = stream.ToArray();
            File.WriteAllBytes(@"d:\测试报告.xlsx", bytes);
        }
    }
}
