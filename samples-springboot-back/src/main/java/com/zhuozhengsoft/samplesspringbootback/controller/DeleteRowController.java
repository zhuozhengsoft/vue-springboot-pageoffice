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
@RequestMapping(value = "/DeleteRow")
public class DeleteRowController {

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务页面
        poCtrl.setServerPage("/api/poserver.zz");
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.addCustomToolButton("删除行", "DeleteRow()", 1);
        //打开Word文档
        poCtrl.webOpen("/api/doc/DeleteRow/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }
}
