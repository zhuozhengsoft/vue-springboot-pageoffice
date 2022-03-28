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
@RequestMapping(value = "/OnWordSelectionChange")
public class OnWordSelectionChangeController {

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        poCtrl.setCustomToolbar(false);
        poCtrl.setJsFunction_OnWordSelectionChange("OnWordSelectionChange()");
        //打开Word文档
        poCtrl.webOpen("/api/doc/OnWordSelectionChange/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


}
