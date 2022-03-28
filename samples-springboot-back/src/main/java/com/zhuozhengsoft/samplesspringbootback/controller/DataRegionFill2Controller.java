package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.awt.*;
import java.util.Map;

@RestController
@RequestMapping(value = "/DataRegionFill2")
public class DataRegionFill2Controller {

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion dataRegion1 = doc.openDataRegion("PO_userName");
        //给数据区域赋值
        dataRegion1.setValue("张三");
        //设置字体样式
        dataRegion1.getFont().setColor(Color.blue);
        dataRegion1.getFont().setSize(24);
        dataRegion1.getFont().setName("隶书");
        dataRegion1.getFont().setBold(true);

        DataRegion dataRegion2 = doc.openDataRegion("PO_deptName");
        dataRegion2.setValue("销售部");
        dataRegion2.getFont().setColor(Color.red);

        poCtrl.setWriter(doc);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);
        //打开Word文档
        poCtrl.webOpen("/api/doc/DataRegionFill2/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


}
