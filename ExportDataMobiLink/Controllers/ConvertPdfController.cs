using GenaratePdf.Projection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GenaratePdf.Controllers
{
    public class ConvertPdfController : Controller
    {
        private readonly GenaratePdfs pdfGenerator;
        public ConvertPdfController()
        {
            pdfGenerator = new GenaratePdfs();
        }
        [HttpPost]
        public ActionResult ReceiveJson(TemplatePdf model)
        {
            model = model ?? new TemplatePdf() { };
            HttpContext.Cache["pdfdata"] = model;
            //return Content(Newtonsoft.Json.JsonConvert.SerializeObject(model));

            return Content("/ConvertPdf/ReceiptPdfFile");
        }

        [HttpGet]
        public ViewResult ReceiptPdfView()
        {
            var order = (TemplatePdf)HttpContext.Cache["pdfdata"];

            return View("~/Views/ConvertPdfs/ProductDetail.cshtml", order);
        }

        [HttpGet]
        public FileContentResult ReceiptPdfFile()
        {
            var local = HttpContext.Request.Url.Authority;

            var url = $"http://{local}/ConvertPdf/ReceiptPdfView";
            //var url =  UnloadingWorkOrderReceiptView(order);

            var pdfByteStream = pdfGenerator.TryRunWkhtml(url);

            Response.AppendHeader("Content-Disposition",
                $"inline; filename = detail.pdf");
            return File(pdfByteStream, "application/pdf");
        }
    }
}