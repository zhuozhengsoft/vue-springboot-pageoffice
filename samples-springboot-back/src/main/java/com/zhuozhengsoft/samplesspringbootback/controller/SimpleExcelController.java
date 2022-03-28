package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.samplesspringbootback.util.GetDirPathUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;



@RestController
@RequestMapping(value = "/SimpleExcel")
public class SimpleExcelController {

        //获取doc目录的磁盘路径
        private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

        @RequestMapping(value ="/Excel")
        @ResponseBody
        public String showExcel(HttpServletRequest request) {

            PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
            poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
            //添加自定义按钮
            poCtrl.addCustomToolButton("保存", "Save", 1);
            poCtrl.addCustomToolButton("关闭", "Close", 21);
            poCtrl.setSaveFilePage("/api/SimpleExcel/save");//设置处理文件保存的请求方法
            //打开word
            poCtrl.webOpen("/api/doc/SimpleExcel/test.xls", OpenModeType.xlsNormalEdit, "张三");
            return poCtrl.getHtmlCode("PageOfficeCtrl1");
        }

        @RequestMapping("save")
        public void save(HttpServletRequest request, HttpServletResponse response) {
            FileSaver fs = new FileSaver(request, response);
            fs.saveToFile(dir + "SimpleExcel/" + fs.getFileName());
            fs.close();
        }
}
