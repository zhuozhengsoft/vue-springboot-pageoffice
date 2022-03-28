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
@RequestMapping(value = "/WordDelBKMK")
public class WordDelBKMKController {

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //隐藏菜单栏
        poCtrl.setMenubar(true);
        //添加自定义按钮
        poCtrl.addCustomToolButton("删除光标处的", "delBookMark()", 7);
        poCtrl.addCustomToolButton("删除选中文本中的", "delChoBookMark()", 7);
        //打开Word文档
        poCtrl.webOpen("/api/doc/WordDelBKMK/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }
}
