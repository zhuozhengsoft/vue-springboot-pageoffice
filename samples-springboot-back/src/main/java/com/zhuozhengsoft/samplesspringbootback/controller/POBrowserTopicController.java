package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.samplesspringbootback.util.GetDirPathUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping(value = "/POBrowserTopic/")
public class POBrowserTopicController {
    String paramValue = "";

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

//    @RequestMapping(value = "index")
//    public ModelAndView showIndex(HttpServletRequest request, Map<String, Object> map, HttpSession session) {
//        String userName = "张三";
//        session.setAttribute("userName", userName);
//        ModelAndView mv = new ModelAndView("POBrowserTopic/index");
//        return mv;
//    }


    @RequestMapping(value = "Word1")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
        //设置保存页面
        poCtrl.setSaveFilePage("/api/POBrowserTopic/save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/api/doc/POBrowserTopic/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping(value = "Word2")
    public Map showWord2(HttpServletRequest request) {
        //获取index页面文本框中的值
        //获取index用？传递过来的id的值
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/doc/POBrowserTopic/test.doc");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/api/doc/POBrowserTopic/test.doc", OpenModeType.docNormalEdit, "张三");
        HashMap<Object, Object> map = new HashMap<>();
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("txt", paramValue);

        return map;
    }

    @RequestMapping(value = "Word3", method = RequestMethod.GET)
    public ModelAndView showWord3(HttpServletRequest request, Map<String, Object> map, HttpSession session) {
        String txt = (String) session.getAttribute("txt");
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存并关闭", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/POBrowserTopic/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("txt", txt);

        ModelAndView mv = new ModelAndView("POBrowserTopic/Word3");
        return mv;
    }


    @RequestMapping("Result2")
    @ResponseBody
    public void Result2(HttpServletRequest request) {
        paramValue = request.getParameter("param");
    }

    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "POBrowserTopic/" + fs.getFileName());
        fs.close();
    }

}
