$(function () {
    var data = {
        listTemplate: [
            {
                productName: "Lê nam kể 6-9 mới ",
                number: 20,
                dVL: "Đôi ",
                cost: 15000,
                discountCost: 0,
                totalCost: 300000
            },
            {
                productName: "Lê nam kể 6-9 mới ",
                number: 20,
                dVL: "Đôi ",
                cost: 15000,
                discountCost: 0,
                totalCost: 300000
            },
            {
                productName: "Lê nam kể 6-9 mới ",
                number: 20,
                dVL: "Đôi ",
                cost: 15000,
                discountCost: 0,
                totalCost: 300000
            },
            {
                productName: "Lê nam kể 6-9 mới ",
                number: 20,
                dVL: "Đôi ",
                cost: 15000,
                discountCost: 0,
                totalCost: 300000
            },
            {
                productName: "Lê nam kể 6-9 mới ",
                number: 20,
                dVL: "Đôi ",
                cost: 15000,
                discountCost: 0,
                totalCost: 300000
            }
        ],
        customer: "Cô Cẩm Nhượng",
        companyName: "Hương Hồng Hạnh",
        companyAddress: "47+49 Hồng Phúc,Ba Đình,Hà Nội",
        companyPhone: "(04)38267927",
        companyFax: "(04)39275068",
        reportDate: "19/03/2017"

    };
    $('#convertpdf').click(function () {
        jQuery.post('/ConvertPdf/ReceiveJson', data, function (response) {
            window.open(response);
            //$("#pdfurl").attr("src", response);
        });
    }
    );

    $('#convertexcel').click(function () {

        jQuery.post('/GenarateExcel/ReceiveJson', data, function (response) {
            window.open(response);
        });
    }
    );

});