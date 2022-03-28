package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;

@RestController
@RequestMapping(value = "/DataRegionFill")
public class DataRegionFillController {

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion dataRegion1 = doc.openDataRegion("PO_userName");
        //给数据区域赋值
        dataRegion1.setValue("张三");

        DataRegion dataRegion2 = doc.openDataRegion("PO_deptName");
        dataRegion2.setValue("销售部");

        poCtrl.setWriter(doc);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);

        //打开Word文档
        poCtrl.webOpen("/api/doc/DataRegionFill/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


}
