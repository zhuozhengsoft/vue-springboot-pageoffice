package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.samplesspringbootback.util.GetDirPathUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

/**
 * @Author: dong
 * @Date: 2022/3/11 16:50
 * @Version 1.0
 */
@RestController
@RequestMapping(value = "/SimpleWord")
public class SimpleWordController {
    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value="/Word")
    @ResponseBody
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("另存为", "SaveAs", 12);
        poCtrl.addCustomToolButton("打印设置", "PrintSet", 0);
        poCtrl.addCustomToolButton("打印", "PrintFile", 6);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
        poCtrl.addCustomToolButton("-", "", 0);
        poCtrl.addCustomToolButton("关闭", "Close", 21);
        poCtrl.setSaveFilePage("/api/SimpleWord/save");//设置保存方法的url
        //打开word
        poCtrl.webOpen("/api/doc/SimpleWord/test.doc", OpenModeType.docNormalEdit, "张三");
        return  poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value ="/Word1")
    @ResponseBody
    public String showWord1(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/SimpleWord/save");//设置处理文件保存的请求方法
        //打开Word文档
        //打开磁盘路径的话有两种写法：1.D:\\doc\\test.doc  2.file://D:/doc/test.doc
        poCtrl.webOpen("file://"+dir+"/SimpleWord/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "SimpleWord/" + fs.getFileName());
        fs.close();
    }


}

