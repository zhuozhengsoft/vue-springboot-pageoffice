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
import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping(value = "/SetHandDrawByUser/")
public class SetHandDrawByUserController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";


    @RequestMapping(value = "Word")
    public Map showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        String userName = request.getParameter("userName");
        if (userName.equals("zhangsan")) userName = "张三";
        if (userName.equals("lisi")) userName = "李四";
        if (userName.equals("wangwu")) userName = "王五";

        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("领导圈阅", "StartHandDraw", 3);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
        poCtrl.setJsFunction_AfterDocumentOpened("ShowByUserName");
        //设置保存页面
        poCtrl.setSaveFilePage("/api/SetHandDrawByUser/save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/api/doc/SetHandDrawByUser/test.doc", OpenModeType.docNormalEdit, "张三");
        HashMap<Object, Object> map = new HashMap<>();
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("userName", userName);
        return map;
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "SetHandDrawByUser/" + fs.getFileName());
        fs.close();
    }

}
