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
@RequestMapping(value = "/FileMakerPDF/")
public class FileMakerPDFController {
    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "OpenPDF")
    public String showindex(HttpServletRequest request) {
        PDFCtrl pdfCtrl1 = new PDFCtrl(request);
        pdfCtrl1.setServerPage("/api/poserver.zz"); //此行必须
        // Create custom toolbar
        pdfCtrl1.addCustomToolButton("打印", "PrintFile()", 6);
        pdfCtrl1.addCustomToolButton("隐藏/显示书签", "SetBookmarks()", 0);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("实际大小", "SetPageReal()", 16);
        pdfCtrl1.addCustomToolButton("适合页面", "SetPageFit()", 17);
        pdfCtrl1.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("首页", "FirstPage()", 8);
        pdfCtrl1.addCustomToolButton("上一页", "PreviousPage()", 9);
        pdfCtrl1.addCustomToolButton("下一页", "NextPage()", 10);
        pdfCtrl1.addCustomToolButton("尾页", "LastPage()", 11);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("向左旋转90度", "SetRotateLeft()", 12);
        pdfCtrl1.addCustomToolButton("向右旋转90度", "SetRotateRight()", 13);
        pdfCtrl1.webOpen("/api/doc/FileMakerPDF/template.pdf");
        return pdfCtrl1.getHtmlCode("PDFCtrl1");
    }

    @RequestMapping(value = "download")
    public void download(HttpServletResponse response) throws IOException {
        String filePath = dir + "FileMakerPDF/" + "template.pdf";
        int fileSize =(int)new File(filePath).length();

        response.reset();
        response.setContentType("application/pdf"); // application/x-excel, application/ms-powerpoint, application/pdf
        response.setHeader("Content-Disposition", "attachment; filename=template.pdf");
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

    @RequestMapping(value = "PDF")
    public String showWord(HttpServletRequest request) {
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage("/api/poserver.zz");
        WordDocument doc = new WordDocument();
        //禁用右击事件
        doc.setDisableWindowRightClick(true);
        //给数据区域赋值，即把数据填充到模板中相应的位置
        doc.openDataRegion("PO_company").setValue("北京卓正志远软件有限公司  ");
        fmCtrl.setSaveFilePage("/api/FileMakerPDF/save?pdfName=template.pdf");
        fmCtrl.setWriter(doc);
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        //fmCtrl.setFileTitle("newfilename.doc");//设置另存为文件的文件名称
        fmCtrl.fillDocumentAsPDF("/api/doc/FileMakerPDF/template.doc", DocumentOpenType.Word, "template.pdf");
        return fmCtrl.getHtmlCode("FileMakerCtrl1");
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        String  pdfName=request.getParameter("pdfName");
        fs.saveToFile(dir + "FileMakerPDF/" + pdfName);
        fs.close();
    }

}
