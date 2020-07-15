using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using System.Collections.Specialized;
using System.Xml.Linq;
using System.Xml;
using System.Runtime.InteropServices;

namespace SendCOT
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("SendCOT")]
    [ComVisible(true)]
    public class WebServiceCOT
    {
        public WebServiceCOT()
        { }
        
        public string ServiceURL { get; set; }
        public string User { get; set; }
        public string Password { get; set; }
        public string ServerName { get; set; }
        public string FileNamePath { get; set; }
        private string FileFormName { get; set; }
        public string rutanuevoarchivo { get; set; }

        public void setParameters(string _URL, string _user, string _pass, string _fileNamePath, string _rutanuevoarchivo)
        {
            ServiceURL = _URL;
            User = _user;
            Password = _pass;
            FileNamePath = _fileNamePath;
            rutanuevoarchivo = _rutanuevoarchivo;
        }       

        public SendCOT.COTReturn SendCOTWS()
        {
            var parms = new NameValueCollection() { { "user", User }, { "password", Password } };
            string sendFileWS = string.Empty;
            string exeption = string.Empty;
            XDocument doc = null;

            SendCOT.COTTBError tbError = null;
            SendCOT.COTTBVoucher tbVoucher = null;
            List<SendCOT.COTError> listErrors = null;
            SendCOT.COTError[] errors = null;

            COTReturn ret = null;

            int i = 0;
            try
            {
                var st = this.UploadFileEx(
                    this.FileNamePath,
                    ServiceURL,
                    FileFormName,
                    "text/plain",
                    parms,
                    User,
                    Password);
                
                System.IO.StreamWriter file = new System.IO.StreamWriter(rutanuevoarchivo);
                file.WriteLine(st.ToString());
                file.Close();

                doc = XDocument.Parse(st);

                tbError = (from r in doc.Descendants("TBError")
                           select new SendCOT.COTTBError
                           {
                               errorCode = r.Element("codigoError").Value,
                               errorMessage = r.Element("mensajeError").Value,
                               errorType = r.Element("tipoError").Value
                           }).FirstOrDefault();

                tbVoucher = (from r in doc.Descendants("TBCOMPROBANTE")
                             select new SendCOT.COTTBVoucher
                             {
                                 companyCuit = r.Element("cuitEmpresa").Value,
                                 voucherNumber = r.Element("numeroComprobante").Value,
                                 fileName = r.Element("nombreArchivo").Value,
                                 integrityCode = r.Element("codigoIntegridad").Value,
                                 uniqueNumber = r.Element("validacionesRemitos").Element("remito").Element("numeroUnico").Value,
                                 processed = r.Element("validacionesRemitos").Element("remito").Element("procesado").Value

                             }).SingleOrDefault();


                listErrors = (from r in doc.Descendants("error")
                              select new SendCOT.COTError
                              {
                                  code = r.Element("codigo").Value,
                                  description = r.Element("descripcion").Value
                              }).ToList();

                errors = new SendCOT.COTError[listErrors.Count];

                foreach (SendCOT.COTError temp in listErrors)
                {
                    errors[i] = temp;
                    i++;
                }

                if (tbVoucher != null)
                    tbVoucher.errors = errors;

                ret = new SendCOT.COTReturn();
                ret.tbError = tbError != null ? tbError : null;
                ret.tbVoucher = tbVoucher != null ? tbVoucher : null;

                return ret;

            }
            catch (Exception ex)
            {
                tbError = new SendCOT.COTTBError();

                tbError.errorMessage = ex.Message;

                ret = new SendCOT.COTReturn();
                ret.tbError = tbError != null ? tbError : null;

                return ret;
            }
        }

        public string tberrorMessage(COTTBError ptbError)
        {
            string ret = "";

            ret = ptbError.errorMessage;

            return ret;
        }

        public string tbVouchererrorMessage(COTTBVoucher pV)
        {
            string ret = "";
            COTError tbError;

            try
            {
                tbError = pV.errors[1];
                ret = tbError.description;
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
                ret = "";
            }
            return ret;
        }


        private string UploadFileEx(string uploadfile, string url, string fileFormName, string contenttype, NameValueCollection querystring, string user, string pass)
        {
            if (string.IsNullOrEmpty(fileFormName))
                fileFormName = "file";
            if ((string.IsNullOrEmpty(contenttype)))
                contenttype = "application/octet-stream";
            Uri uri = new Uri(url);
            string boundary = "----------" + DateTime.Now.Ticks.ToString("x");
            HttpWebRequest webrequest = (HttpWebRequest)WebRequest.Create(uri);
            
            webrequest.ContentType = "multipart/form-data; boundary=" + boundary;
            webrequest.Method = "POST";

            StringBuilder sb = new StringBuilder();
            if (querystring != null)
            {
                foreach (string key in querystring.Keys)
                {
                    sb.AppendFormat("--{0}\r\n", boundary);
                    sb.AppendFormat("Content-Disposition:  form-data; name=\"{0}\";\r\n\r\n{1}\r\n", key, querystring[key]);
                }
            }
            sb.Append("--");
            sb.Append(boundary);
            sb.Append("\r\n");
            sb.Append("Content-Disposition: form-data; name=\"");
            sb.Append(fileFormName);
            sb.Append("\"; filename=\"");
            sb.Append(Path.GetFileName(uploadfile));
            sb.Append("\"");
            sb.Append("\r\n");
            sb.Append("Content-Type: ");
            sb.Append(contenttype);
            sb.Append("\r\n");
            sb.Append("\r\n");

            string postHeader = sb.ToString();
            byte[] postHeaderBytes = Encoding.UTF8.GetBytes(postHeader);            
            byte[] boundaryBytes = Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");

            FileStream fileStream = new FileStream(uploadfile, FileMode.Open, FileAccess.Read);
            long length = postHeaderBytes.Length + fileStream.Length + boundaryBytes.Length;
            webrequest.ContentLength = length;

            Stream requestStream = webrequest.GetRequestStream();           
            requestStream.Write(postHeaderBytes, 0, postHeaderBytes.Length);
            
            byte[] buffer = new Byte[checked((uint)Math.Min(4096, (int)fileStream.Length))];
            int bytesRead = 0;
            while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) != 0)
                requestStream.Write(buffer, 0, bytesRead);
           
            requestStream.Write(boundaryBytes, 0, boundaryBytes.Length);
            WebResponse responce = webrequest.GetResponse();
            Stream s = responce.GetResponseStream();
            string ret;
            using (StreamReader sr = new StreamReader(s, System.Text.Encoding.UTF7))
                ret = sr.ReadToEnd();

            fileStream.Close();
            requestStream.Close();
            responce.Close();

            return ret;
        }
    }

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("COTReturn")]
    [ComVisible(true)]
    public class COTReturn
    {
        public SendCOT.COTTBError tbError { get; set; }
        public SendCOT.COTTBVoucher tbVoucher { get; set; }
    }

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("COTTBError")]
    [ComVisible(true)]
    public class COTTBError
    {
        public string errorType { get; set; }
        public string errorCode { get; set; }
        public string errorMessage { get; set; }

    }

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("COTTBVoucher")]
    [ComVisible(true)]
    public class COTTBVoucher
    {
        public string companyCuit { get; set; }
        public string voucherNumber { get; set; }
        public string fileName { get; set; }
        public string integrityCode { get; set; }
        public string uniqueNumber { get; set; }
        public string processed { get; set; }

        public SendCOT.COTError[] errors { get; set; }
    }

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("COTError")]
    [ComVisible(true)]
    public class COTError
    {
        public string code { get; set; }
        public string description { get; set; }
    }
}
