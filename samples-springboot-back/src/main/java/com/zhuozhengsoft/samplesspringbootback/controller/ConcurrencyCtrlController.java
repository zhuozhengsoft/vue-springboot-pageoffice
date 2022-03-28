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
@RequestMapping(value = "/ConcurrencyCtrl")
public class ConcurrencyCtrlController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";


    @RequestMapping(value = "/Word")
    public Map<Object,Object> showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        String userName = "somebody";
        String userId = request.getParameter("userid").toString();
        if (userId.equals("1")) {
            userName = "张三";
        } else {
            userName = "李四";
        }
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置并发控制时间
        poCtrl.setTimeSlice(20);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/ConcurrencyCtrl/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/ConcurrencyCtrl/test.doc", OpenModeType.docRevisionOnly, userName);
        Map<Object, Object> map = new HashMap<>();
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("userName", userName);
        return map;
    }

    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "ConcurrencyCtrl/" + fs.getFileName());
        fs.close();
    }

}
