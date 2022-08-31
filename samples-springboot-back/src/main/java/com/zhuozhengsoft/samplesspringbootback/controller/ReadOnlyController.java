package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@RestController
@RequestMapping(value = "/ReadOnly")
public class ReadOnlyController {

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
		WordDocument wordDoc=new WordDocument();
        wordDoc.setDisableWindowSelection(true);//禁止选中
        wordDoc.setDisableWindowRightClick(true);//禁止右键

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        poCtrl.setMenubar(false);//隐藏菜单栏
        poCtrl.setOfficeToolbars(false);//隐藏Office工具条
        poCtrl.setCustomToolbar(false);//隐藏自定义工具栏
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");
        //设置页面的显示标题
        poCtrl.setCaption("演示：文件在线安全浏览");
        //打开Word文档
	    poCtrl.setWriter(wordDoc);//此句必须
        poCtrl.webOpen("/api/doc/ReadOnly/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }
}
