package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@RestController
@RequestMapping(value = "/ExcelInsertImage")
public class ExcelInsertImageController {

    @RequestMapping(value = "/Excel")
    public String showExcel(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        Workbook workBook = new Workbook();
        Sheet sheet1 = workBook.openSheet("Sheet1");
        sheet1.openCell("A1").setValue("[image]/api/doc/ExcelInsertImage/image/logo.jpg[/image]");

        poCtrl.setWriter(workBook);//此行必须

        //打开Word文档
        poCtrl.webOpen("/api/doc/ExcelInsertImage/test.xls", OpenModeType.xlsNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

}
