package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.PDFCtrl;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@RestController
@RequestMapping(value = "/PDFSearch")
public class PDFSearchController {

    @RequestMapping(value = "/PDF")
    public String showPDF(HttpServletRequest request) {
        PDFCtrl poCtrl = new PDFCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz"); //此行必须

        // Create custom toolbar
        poCtrl.addCustomToolButton("搜索", "SearchText()", 0);
        poCtrl.addCustomToolButton("搜索下一个", "SearchTextNext()", 0);
        poCtrl.addCustomToolButton("搜索上一个", "SearchTextPrev()", 0);
        poCtrl.addCustomToolButton("实际大小", "SetPageReal()", 16);
        poCtrl.addCustomToolButton("适合页面", "SetPageFit()", 17);
        poCtrl.addCustomToolButton("适合宽度", "SetPageWidth()", 18);

        //打开Word文档
        poCtrl.webOpen("/api/doc/PDFSearch/test.pdf");
        return poCtrl.getHtmlCode("PDFCtrl1");
    }

}
