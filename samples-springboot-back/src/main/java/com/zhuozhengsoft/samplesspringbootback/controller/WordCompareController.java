package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@RestController
@RequestMapping(value = "/WordCompare")
public class WordCompareController {

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        poCtrl.addCustomToolButton("显示A文档", "ShowFile1View()", 0);
        poCtrl.addCustomToolButton("显示B文档", "ShowFile2View()", 0);
        poCtrl.addCustomToolButton("显示比较结果", "ShowCompareView()", 0);
        //poCtrl.wordCompare("doc/aaa1.doc", "doc/aaa2.doc", OpenModeType.docReadOnly, "张三");
        poCtrl.wordCompare("/api/doc/WordCompare/aaa1.doc", "/api/doc/WordCompare/aaa2.doc", OpenModeType.docAdmin, "张三");

        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


}
