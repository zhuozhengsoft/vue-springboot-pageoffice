package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.samplesspringbootback.util.GetDirPathUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@RestController
@RequestMapping(value = "/SaveFirstPageAsImg")
public class SaveFirstPageAsImgController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("保存首页为图片", "SaveFirstAsImg()", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/SaveFirstPageAsImg/save");//设置处理文件保存的请求方法


        //打开Word文档
        poCtrl.webOpen("/api/doc/SaveFirstPageAsImg/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        if (fs.getFileExtName().equals(".jpg")) {
            fs.saveToFile(dir + "SaveFirstPageAsImg/images/" + fs.getFileName());
        } else {
            fs.saveToFile(dir + "SaveFirstPageAsImg/" + fs.getFileName());
        }
        fs.setCustomSaveResult("ok");

        fs.close();
    }

}
