using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography.Xml;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using ExcelDataReader.Log;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace TestApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class Product : Controller
    {
        [HttpGet("GetXml")]
        public string GetXml()
        {
            try
            {
                // hata almak için exception oluşturduk.
                throw new Exception();

                XDocument document = XDocument.Load(@"..\src\datapool\data.xml"); 
                  
                foreach (var attr in document.Descendants())
                {
                    attr.RemoveAttributes();
                }

                return document.ToString(); 
            }
            catch (Exception ex)
            {
                //  hata alma durumunda loglama yapıyoruz  kütüphanemiz => using ExcelDataReader.Log;
                LogManager.Log(this).Debug("GetXml Exception {0}", ex.Message);
            }
            return null;
        }

        [HttpGet("GetJson")]
        public string Get()
        {
            try
            {
                XDocument document = XDocument.Load(@"..\src\datapool\data.xml"); 

                foreach (var attr in document.Descendants())
                {
                    attr.RemoveAttributes();
                }

                return JsonConvert.SerializeXNode(document);  
            }
            catch (Exception ex)
            {
                //  hata alma durumunda loglama yapıyoruz kütüphanemiz => using ExcelDataReader.Log;
                LogManager.Log(this).Debug("GetJson Exception {0}", ex.Message);
            }
            return null;
        }

    }
}
