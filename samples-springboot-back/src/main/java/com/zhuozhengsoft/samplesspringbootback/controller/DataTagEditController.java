package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import com.zhuozhengsoft.samplesspringbootback.util.GetDirPathUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@RestController
@RequestMapping(value = "/DataTagEdit")
public class DataTagEditController {
    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        WordDocument doc = new WordDocument();
        doc.getTemplate().defineDataTag("{ 甲方 }");
        doc.getTemplate().defineDataTag("{ 乙方 }");
        doc.getTemplate().defineDataTag("{ 担保人 }");
        doc.getTemplate().defineDataTag("【 合同日期 】");
        doc.getTemplate().defineDataTag("【 合同编号 】");

        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("定义数据标签", "ShowDefineDataTags()", 20);

        poCtrl.setSaveFilePage("/api/DataTagEdit/save");

        poCtrl.setTheme(ThemeType.Office2007);
        poCtrl.setBorderStyle(BorderStyleType.BorderThin);
        poCtrl.setWriter(doc);
        //打开Word文档
        poCtrl.webOpen("/api/doc/DataTagEdit/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "DataTagDlg")
    public String dataTagDlg(HttpServletRequest request) {
        return "DataTagEdit/DataTagDlg";
    }
    
    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "DataTagEdit/" + fs.getFileName());
        fs.close();
    }

}
