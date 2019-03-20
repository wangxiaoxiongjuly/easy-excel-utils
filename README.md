# easy-excel-utils
基于阿里easyExcel包装的工具类，poi版本3.17

##模块介绍
- easy-excel-util #工具类依赖#
- my-base-common #测试使用的其他依赖#
- my-web #测试使用的web工程#

##启动介绍


my-web基于SpringBoot ，默认不使用外置tomcat容器，使用外置容器时将WebServletInitializerConf的注释打开，并且将打包方式修改为war


调用AdminApplication中的Main方法启动内部tomcat容器

##导出代码位置及调用示例
my-web中 EasyExcelUtilController 

`
ExcelUtil.writeExcel(response,list,"导出测试","没有设定sheet名称", ExcelTypeEnum.XLSX,ExportTestModel.class);
`

##导入
my-web中 EasyExcelUtilController 

`
List<ExportTestModel> list = ExcelUtil.readExcel(files.get(0),ExportTestModel.class);
`

