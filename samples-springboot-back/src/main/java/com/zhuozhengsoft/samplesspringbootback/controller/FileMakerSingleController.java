package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import com.zhuozhengsoft.samplesspringbootback.util.GetDirPathUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;

@RestController
@RequestMapping(value = "/FileMakerSingle/")
public class FileMakerSingleController {
    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "OpenWord")
    public String showindex(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setCustomToolbar(false);
        poCtrl.setServerPage("/api/poserver.zz");
        poCtrl.webOpen("/api/doc/FileMakerSingle/maker.doc", OpenModeType.docReadOnly, "张佚名");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "download")
    public void download(HttpServletResponse response) throws IOException {
        String filePath = dir + "FileMakerSingle/" + "maker.doc";
        int fileSize =(int)new File(filePath).length();

        response.reset();
        response.setContentType("application/msword"); // application/x-excel, application/ms-powerpoint, application/pdf
        response.setHeader("Content-Disposition", "attachment; filename=maker.doc");
        response.setContentLength(fileSize);

        OutputStream outputStream = response.getOutputStream();
        InputStream inputStream = new FileInputStream(filePath);
        byte[] buffer = new byte[10240];
        int i = -1;
        while ((i = inputStream.read(buffer)) != -1) {
            outputStream.write(buffer, 0, i);
        }

        outputStream.flush();
        outputStream.close();
        inputStream.close();
    }

    @RequestMapping(value = "Word")
    public String showWord(HttpServletRequest request) {
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage("/api/poserver.zz");
        WordDocument doc = new WordDocument();
        //禁用右击事件
        doc.setDisableWindowRightClick(true);
        //给数据区域赋值，即把数据填充到模板中相应的位置
        doc.openDataRegion("PO_company").setValue("北京卓正志远软件有限公司  ");
        fmCtrl.setSaveFilePage("/api/FileMakerSingle/save");
        fmCtrl.setWriter(doc);
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        //fmCtrl.setFileTitle("newfilename.doc");//设置另存为文件的文件名称
        fmCtrl.fillDocument("/api/doc/FileMakerSingle/test.doc", DocumentOpenType.Word);
        return fmCtrl.getHtmlCode("FileMakerCtrl1");
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "FileMakerSingle/" + "maker.doc");
        fs.close();
    }

}
