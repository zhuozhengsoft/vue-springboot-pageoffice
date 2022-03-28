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
import java.util.Map;

@RestController
@RequestMapping(value = "/InsertImageSetSize")
public class InsertImageSetSizeController {

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        WordDocument worddoc = new WordDocument();
        //先在要插入word文件的位置手动插入书签,书签必须以“PO_”为前缀
        //给DataRegion赋值,值的形式为："[word]word文件路径[/word]、[excel]excel文件路径[/excel]、[image]图片路径[/image]"
        DataRegion data1 = worddoc.openDataRegion("PO_p1");
        // [image]标签中的width和height属性的单位是：英磅
        data1.setValue("[image width=200.2 height=200]/api/doc/InsertImageSetSize/1.jpg[/image]");
        poCtrl.setWriter(worddoc);
        poCtrl.setCaption("演示：后台编程插入图片到数据区域并设置图片大小(企业版)");
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏自定义工具栏
        poCtrl.setCustomToolbar(false);

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertImageSetSize/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


}
