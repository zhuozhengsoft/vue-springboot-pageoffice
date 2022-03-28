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
@RequestMapping(value = "/ControlBars")
public class ControlBarsController {

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        poCtrl.setTitlebar(false); //隐藏标题栏
        poCtrl.setMenubar(false); //隐藏菜单栏
        poCtrl.setOfficeToolbars(false);//隐藏Office工具条
        poCtrl.setCustomToolbar(false);//隐藏自定义工具栏
        //打开Word文档
        poCtrl.webOpen("/api/doc/ControlBars/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

}
