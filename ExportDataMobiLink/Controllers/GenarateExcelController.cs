using GenaratePdf.ExcelFiles;
using GenaratePdf.Projection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace GenaratePdf.Controllers
{
    public class GenarateExcelController : Controller
    {
        // GET: GenarateExcel
        //public ActionResult Index()
        //{
        //    return View();
        //}

        //public void ExportToExcel(TemplatePdf model)
        //{
        //    string Filename = "ExcelFrom.xls";
        //    string FolderPath = HttpContext.Server.MapPath("/ExcelFiles/");
        //    string FilePath = Path.Combine(FolderPath, Filename);


        //    //Step-1: Checking: the file name exist in server, if it is found then remove from server.------------------
        //    if (System.IO.File.Exists(FilePath))
        //    {
        //        System.IO.File.Delete(FilePath);
        //    }
        //    //var order = (TemplatePdf)HttpContext.Cache["pdfdata"];
        //    //Step-2: Get Html Data & Converted to String----------------------------------------------------------------
        //    string HtmlResult = RenderRazorViewToString("~/Views/GenarateExcel/ProductDetailExcel.cshtml", model);

        //    //Step-4: Html Result store in Byte[] array------------------------------------------------------------------
        //    byte[] ExcelBytes = Encoding.ASCII.GetBytes(HtmlResult);

        //    //Step-5: byte[] array converted to file Stream and save in Server------------------------------------------- 
        //    using (Stream file = System.IO.File.OpenWrite(FilePath))
        //    {
        //        file.Write(ExcelBytes, 0, ExcelBytes.Length);
        //    }

        //    //Step-6: Download Excel file 
        //    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //    Response.AddHeader("Content-Disposition", "attachment; filename=" + Path.GetFileName(Filename));
        //    Response.WriteFile(FilePath);
        //    Response.Flush();
        //    Response.End();

        //}

        //protected string RenderRazorViewToString(string viewName, object model)
        //{
        //    if (model != null)
        //    {
        //        ViewData.Model = model;
        //    }
        //    using (StringWriter sw = new StringWriter())
        //    {
        //        ViewEngineResult viewResult = ViewEngines.Engines.FindPartialView(ControllerContext, viewName);
        //        ViewContext viewContext = new ViewContext(ControllerContext, viewResult.View, ViewData, TempData, sw);
        //        viewResult.View.Render(viewContext, sw);
        //        viewResult.ViewEngine.ReleaseView(ControllerContext, viewResult.View);
        //        return sw.GetStringBuilder().ToString();
        //    }
        //}
        //[HttpPost]
        //public ActionResult ExcelReceiveJson(TemplatePdf model)
        //{
        //    model = model ?? new TemplatePdf() { };
        //    model.PoNumber = "xxx";
        //    HttpContext.Cache["exceldata"] = model;
        //    //return Content(Newtonsoft.Json.JsonConvert.SerializeObject(model));

        //    return Content("/GenarateExcel/ProductDetail");
        //}
        [HttpPost]
        public ActionResult ReceiveJson(TemplatePdf model)
        {
            model = model ?? new TemplatePdf() { };
            HttpContext.Cache["exceldata"] = model;
            //return Content(Newtonsoft.Json.JsonConvert.SerializeObject(model));

            return Content("/GenarateExcel/ProductDetail");
        }

        [HttpGet]
        public FileContentResult ProductDetail()
        {
            var model = (TemplatePdf)HttpContext.Cache["exceldata"];
            string Filename = "Mobilink" + DateTime.Now.ToString("mm_dd_yyy_hh_ss_tt") + ".xlsx";
            string FolderPath = HttpContext.Server.MapPath("/ExcelFiles/");
            string FilePath = Path.Combine(FolderPath, Filename);
            var excelGenerator = new GenarateReportExcel();

            var fileBytes = excelGenerator.GenerateReport(model);

            return XlsxFile(fileBytes, Filename);

            //using (Stream file = System.IO.File.OpenWrite(FilePath))
            //{
            //    file.Write(fileBytes, 0, fileBytes.Length);
            //}
            ////Step - 6: Download Excel file
            //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            ////Response.ContentType = "application/x-download";
            //Response.ContentEncoding = System.Text.Encoding.UTF8;
            //Response.AddHeader("Content-Disposition", "attachment; filename=" + Path.GetFileName(Filename));
            //Response.WriteFile(FilePath);
            //Response.Flush();
            //Response.End();

        }
        protected FileContentResult XlsxFile(byte[] contents, string downloadName)
        {
            return File(contents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", downloadName);
        }
    }
}