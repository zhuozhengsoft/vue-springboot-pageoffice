package com.zhuozhengsoft.samplesspringbootback.controller;

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
import java.io.UnsupportedEncodingException;
import java.util.Map;

@RestController
@RequestMapping(value = "/ExtractImage")
public class ExtractImageController {
    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        poCtrl.addCustomToolButton("保存图片", "Save", 1);
        WordDocument wordDoc = new WordDocument();
        //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
        DataRegion dataRegion1 = wordDoc.openDataRegion("PO_image");
        dataRegion1.setEditing(true);//放图片的数据区域是可以编辑的，其它部分不可编辑
        poCtrl.setWriter(wordDoc);
        //设置保存页面
        poCtrl.setSaveDataPage("/api/ExtractImage/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/ExtractImage/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws UnsupportedEncodingException {
        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
        com.zhuozhengsoft.pageoffice.wordreader.DataRegion dr = doc.openDataRegion("PO_image");
        //将提取的图片保存到服务器上，图片的名称为:a.jpg
        dr.openShape(1).saveAsJPG(dir + "ExtractImage/" + "a.jpg");

        doc.setCustomSaveResult(java.net.URLEncoder.encode("保存成功,文件保存到：","utf-8")+ dir + "ExtractImage/" + "/a.jpg");
        doc.close();
    }

}
