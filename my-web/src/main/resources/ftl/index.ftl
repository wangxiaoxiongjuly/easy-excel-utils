<!DOCTYPE html>
<html >
<#include "base/baseStyle.ftl">
<body>
<H1>Hello,World!</H1>
<div>
    <a href="/easyExcelUtil/getExportData" target="_blank">EasyExcel导出测试</a>
</div>
<div style="position: relative;">
    <input style="width: 68px; position: absolute;opacity: 0;" id="importExcel" type="file" accept=".xlsx">
    <button >导入测试</button>
</div>


<#include "base/baseJs.ftl">
<script src="/assets/jqueryFileUpload-9.10.0/vendor/jquery.ui.widget.js?1"></script>
<script src="/assets/jqueryFileUpload-9.10.0/jquery.fileupload.js?1"></script>
<script>
    $('#importExcel').fileupload({
        url: "/easyExcelUtil/importExcel",
        dataType: 'json',
        success: function (data) {
            layer.msg(data.msg);
        }
    });
</script>
</body>
</html>