package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import com.zhuozhengsoft.samplesspringbootback.util.GetDirPathUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.util.Map;

@RestController
@RequestMapping(value = "/ClickDataRegion")
public class ClickDataRegionController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        WordDocument doc = new WordDocument();
        DataRegion dataReg = doc.openDataRegion("PO_deptName");
        dataReg.getShading().setBackgroundPatternColor(Color.pink);
        //dataReg.setEditing(true);
        poCtrl.setWriter(doc);
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.setJsFunction_OnWordDataRegionClick("OnWordDataRegionClick()");
        poCtrl.setOfficeToolbars(false);
        poCtrl.setCaption("为方便用户知道哪些地方可以编辑，所以设置了数据区域的背景色");
        //设置保存页面
        poCtrl.setSaveFilePage("/api/ClickDataRegion/save");//设置处理文件保存的请求方法


        //打开Word文档
        poCtrl.webOpen("/api/doc/ClickDataRegion/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "HTMLPage")
    public String hTMLPage() {

        return "hTMLPage";
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "ClickDataRegion/" + fs.getFileName());
        fs.close();
    }

}
