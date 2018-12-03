using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GenaratePdf.Projection
{
    public class TemplatePdf
    {
        public IList<Template> ListTemplate { get; set; }
        public string Customer { get; set; }
        public string CompanyName { get; set; }
        public string CompanyAddress { get; set; }
        public string CompanyPhone { get; set; }
        public string CompanyFax { get; set; }
        public string ReportDate { get; set; }
    }

    public class Template
    {
        public string ProductName { get; set; }

        public int Number { get; set; }

        public string DVL { get; set; }

        public int Cost { get; set; }

        public int DiscountCost { get; set; }

        public int TotalCost { get; set; }

    }
}