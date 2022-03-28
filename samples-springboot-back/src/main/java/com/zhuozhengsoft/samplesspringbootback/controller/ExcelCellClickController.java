package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Table;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.Map;

@RestController
@RequestMapping(value = "/ExcelCellClick")
public class ExcelCellClickController {

    @RequestMapping(value = "/Excel")
    public String showExcel(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");

        //定义table对象，设置table对象的设置范围
        Table table = sheet.openTable("B4:D8");
        //设置table对象的提交名称，以便保存页面获取提交的数据
        table.setSubmitName("Info");

        // 设置响应单元格点击事件的js function
        poCtrl.setJsFunction_OnExcelCellClick("OnCellClick()");

        poCtrl.setWriter(workBook);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);

        //设置保存页面
        poCtrl.setSaveDataPage("/api/ExcelCellClick/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/ExcelCellClick/test.xls", OpenModeType.xlsSubmitForm, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response, Map<String, Object> map) throws IOException {
        response.setCharacterEncoding("utf-8");//解决返回的数据中文乱码问题
        com.zhuozhengsoft.pageoffice.excelreader.Workbook workBook = new com.zhuozhengsoft.pageoffice.excelreader.Workbook(request, response);
        com.zhuozhengsoft.pageoffice.excelreader.Sheet sheet = workBook.openSheet("Sheet1");
        com.zhuozhengsoft.pageoffice.excelreader.Table table = sheet.openTable("B4:D8");
        String content = "";
        int result = 0;
        while (!table.getEOF()) {

            //获取提交的数值
            //DataFields.Count标识的是table的列数
            if (!table.getDataFields().getIsEmpty()) {
                content += "<br/>月份名称：" + table.getDataFields().get(0).getText();
                content += "<br/>计划完成量：" + table.getDataFields().get(1).getText();
                content += "<br/>实际完成量：" + table.getDataFields().get(2).getText();

                content += "<br/>*********************************************";
            }
            //循环进入下一行
            table.nextRow();
        }
        table.close();

        workBook.showPage(500, 400);
        workBook.close();
        response.setContentType("text/plain; charset=utf-8");
        response.getWriter().write(content);

    }


}
