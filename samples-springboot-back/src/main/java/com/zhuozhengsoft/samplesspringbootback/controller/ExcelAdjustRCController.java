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
@RequestMapping(value = "/ExcelAdjustRC")
public class ExcelAdjustRCController {

    @RequestMapping(value = "/Excel")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        poCtrl.setCustomToolbar(false);
        Workbook wb = new Workbook();
        Sheet sheet1 = wb.openSheet("Sheet1");
        //设置当工作表只读时，是否允许用户手动调整行列。
        sheet1.setAllowAdjustRC(true);
        poCtrl.setWriter(wb);//此行必须


        //打开Word文档
        poCtrl.webOpen("/api/doc/ExcelAdjustRC/test.xls", OpenModeType.xlsReadOnly, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


}
