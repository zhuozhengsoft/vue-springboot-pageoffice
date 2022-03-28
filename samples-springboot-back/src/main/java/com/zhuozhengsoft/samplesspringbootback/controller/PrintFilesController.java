package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.DocumentOpenType;
import com.zhuozhengsoft.pageoffice.FileMakerCtrl;
import com.zhuozhengsoft.pageoffice.FileSaver;
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
@RequestMapping(value = "/PrintFiles/")
public class PrintFilesController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "Word")
    public String showWord(HttpServletRequest request) {
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage("/api/poserver.zz");
        String id = request.getParameter("id");
        if (id != null && id.length() > 0) {
            WordDocument doc = new WordDocument();
            //给数据区域赋值，即把数据填充到模板中相应的位置
            doc.openDataRegion("PO_company").setValue("北京卓正志远软件有限公司  " + id);
            fmCtrl.setSaveFilePage("/api/PrintFiles/save?id=" + id);
            fmCtrl.setWriter(doc);
            fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
            fmCtrl.fillDocument("/api/doc/PrintFiles/test.doc", DocumentOpenType.Word);
        }
        return fmCtrl.getHtmlCode("FileMakerCtrl1");
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        String id = request.getParameter("id");
        String err = "";
        if (id != null && id.length() > 0) {
            String fileName = "/maker" + id + fs.getFileExtName();
            fs.saveToFile(dir + "PrintFiles/" + fileName);
        } else {
            throw new RuntimeException("id不正常");
        }

        fs.close();
    }

}
