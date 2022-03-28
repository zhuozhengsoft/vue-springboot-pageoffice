package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.samplesspringbootback.util.GetDirPathUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@RestController
@RequestMapping(value = "/FileMakerConvertPDFs/")
public class FileMakerConvertPDFsController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "index")
    public String showindex() {
        return dir + "FileMakerConvertPDFs";
    }

    @RequestMapping(value = "Word")
    public String showWord(HttpServletRequest request) {

        String path = request.getContextPath();
        String filePath = "";
        String id = request.getParameter("id").trim();
        System.out.println(id);
        if ("1".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "PageOffice产品简介.doc";
        }
        if ("2".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "Pageoffice客户端安装步骤.doc";
        }
        if ("3".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "PageOffice的应用领域.doc";
        }
        if ("4".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "PageOffice产品对客户端环境要求.doc";
        }

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/FileMakerConvertPDFs/save");//设置处理文件保存的请求方法
        filePath = filePath.replace("/", "\\");
        //打开Word文档
        poCtrl.webOpen(filePath, OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "Convert")
    public String showWConvert(HttpServletRequest request) {
        String path = request.getContextPath();
        String filePath = "";
        String id = request.getParameter("id").trim();
        if ("1".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "PageOffice产品简介.doc";
        }
        if ("2".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "Pageoffice客户端安装步骤.doc";
        }
        if ("3".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "PageOffice的应用领域.doc";
        }
        if ("4".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "PageOffice产品对客户端环境要求.doc";
        }

        filePath = filePath.replace("/", "\\");
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage("/api/poserver.zz");
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        fmCtrl.setSaveFilePage("/api/FileMakerConvertPDFs/save");
        fmCtrl.fillDocumentAsPDF(filePath, DocumentOpenType.Word, "a.pdf");
        return fmCtrl.getHtmlCode("FileMakerCtrl1");
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "FileMakerConvertPDFs/" + fs.getFileName());
        fs.close();
    }

}
