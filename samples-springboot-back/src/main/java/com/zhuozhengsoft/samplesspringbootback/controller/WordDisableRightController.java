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
@RequestMapping(value = "/WordDisableRight")
public class WordDisableRightController {

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        WordDocument wordDoc = new WordDocument();
        wordDoc.setDisableWindowRightClick(true);//禁止word鼠标右键
        //wordDoc.setDisableWindowDoubleClick(true);//禁止word鼠标双击
        //wordDoc.setDisableWindowSelection(true);//禁止在word中选择文件内容
        poCtrl.setWriter(wordDoc);
        //打开Word文档
        poCtrl.webOpen("/api/doc/WordDisableRight/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


}
