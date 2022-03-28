package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping(value = "/Token")
public class Token {

    @RequestMapping(value = "/Word")
    public Map<String ,String> showWord(HttpServletRequest request) {
        Map<String, String> map = new HashMap<>();
        //获取token值
        String mytoken=request.getHeader("Authorization");
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("关闭", "Close", 21);
        //打开Word文档
        poCtrl.webOpen("/api/doc/CallParentFunction/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("mytoken",mytoken);
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        return map;
    }


}
