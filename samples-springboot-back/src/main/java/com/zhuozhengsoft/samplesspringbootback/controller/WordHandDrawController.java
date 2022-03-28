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
@RequestMapping(value = "/WordHandDraw")
public class WordHandDrawController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "/Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("开始手写", "StartHandDraw()", 5);
        poCtrl.addCustomToolButton("设置线宽", "SetPenWidth()", 5);
        poCtrl.addCustomToolButton("设置颜色", "SetPenColor()", 5);
        poCtrl.addCustomToolButton("设置笔型", "SetPenType()", 5);
        poCtrl.addCustomToolButton("设置缩放", "SetPenZoom()", 5);
        poCtrl.addCustomToolButton("访问手写集", "GetHandDrawList()", 6);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/WordHandDraw/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/WordHandDraw/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "WordHandDraw/" + fs.getFileName());
        fs.close();
    }

}
