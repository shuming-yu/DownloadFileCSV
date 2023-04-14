<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="DownloadFileCSV.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <!-- jQuery 3.x -->
    <script src="https://code.jquery.com/jquery-3.6.4.min.js" integrity="sha256-oP6HI9z1XaZNBrJURtCoUT5SUnxFr8s3BzRl+cbzUq8=" crossorigin="anonymous"></script>
    <!-- moment.js -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.4/moment.min.js"></script>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
        </div>
    </form>
</body>
</html>

<script>
    
    /**
     * 將查詢盤點紀錄匯出為csv
     */
    function Download() {
        var plantCode = $("#ddlPlantCode").val();
        var storageLoc = $("#ddlStorageLoc").val();
        var partNoList = $("#txtPartNo").val();
        var locations = $("#txtLocation").val();
        var editor = $("#ddlEditor").val();
        var startCheckDate = $("#startDate").val();
        var endCheckDate = $("#endDate").val();

        //ShowLoaingIMG();
        var url = routeURI + String.format("GetExportCheckReportToCSV?plantCode={0}&storageLoc={1}&partNoList={2}&locations={3}&editor={4}&startCheckDate={5}&endCheckDate={6}&containHeader={7}",
            plantCode, storageLoc, partNoList, locations, editor, startCheckDate, endCheckDate, true);
        $.ajax({
            type: "GET",
            url: url,
            async: true,
            cache: false,
            headers: {
                'Language': "zh-TW"
            },
            contentType: "application/octet-stream; charset=utf-8",
            success: function (blob, status, xhr) {
                //HideLoaingIMG();
                var filename = "CheckReport-" + moment().format('YYYYMMDDHHmm');
                ExportCSV(blob, filename);
            },
            error: function (xhr, status, error) {
                //HideLoaingIMG();
                var msg = status == 'timeout' ? 'Error: Timeout' : xhr.responseText;
                //ShowAllMessage(msg);
            }
        });
    }

    function ExportCSV(textTemplate, fileName) {
        if (!fileName) {
            fileName = "data.csv";
        }
        else if (!/\w+(\.\w+)$/.test(fileName)) {
            fileName += '.csv';
        }

        var ua = window.navigator.userAgent;
        var msie = ua.indexOf("MSIE ");

        if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./)) {
            // If Internet Explorer
            var frameId = excelFrame;
            frameId.document.open("text/csv", "replace");
            //使用逗號分隔 "sep=,"
            frameId.document.write('sep=,\r\n' + textTemplate);
            frameId.document.close();
            frameId.focus();
            return frameId.document.execCommand("SaveAs", true, fileName);
        }
        else {
            //other browser not tested on IE 11
            var blob = new Blob([textTemplate], { type: 'application/octet-stream;charset=utf-8' });
            var windowURL = window.URL || window.webkitURL;
            var path = windowURL.createObjectURL(blob);
            var link = $('<a>').attr('href', path)
                .attr('download', fileName)
                .attr('target', '_blank')
                .css({ display: 'none' })
                .appendTo($('body'));

            link[0].click();
            link.remove();
            //return window.open('data:text/csv,' + encodeURIComponent(textTemplate));
        }
    }

</script>