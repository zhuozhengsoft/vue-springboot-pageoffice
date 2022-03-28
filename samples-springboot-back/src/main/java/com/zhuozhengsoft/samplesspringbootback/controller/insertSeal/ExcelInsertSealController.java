package com.zhuozhengsoft.samplesspringbootback.controller.insertSeal;

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
@RequestMapping(value = "/InsertSeal/Excel/")
public class ExcelInsertSealController {
    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "AddSeal/Excel1")
    public String showExcel(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Excel/AddSeal/save");//设置处理文件保存的请求方法

        //打开Excel文档
        poCtrl.webOpen("/api/doc/InsertSeal/Excel/AddSeal/test1.xls", OpenModeType.xlsNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping(value = "AddSeal/Excel2")
    public String showExcel2(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Excel/AddSeal/save");//设置处理文件保存的请求方法

        //打开Excel文档
        poCtrl.webOpen("/api/doc/InsertSeal/Excel/AddSeal/test2.xls", OpenModeType.xlsNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSeal/Excel3")
    public String showExcel3(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Excel/AddSeal/save");//设置处理文件保存的请求方法

        //打开Excel文档
        poCtrl.webOpen("/api/doc/InsertSeal/Excel/AddSeal/test3.xls", OpenModeType.xlsNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSeal/Excel4")
    public String showExcel4(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("验证文档", "VerifySeal()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Excel/AddSeal/save");//设置处理文件保存的请求方法

        //打开Excel文档
        poCtrl.webOpen("/api/doc/InsertSeal/Excel/AddSeal/test4.xls", OpenModeType.xlsNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSeal/Excel5")
    public String showExcel5(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除指定印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("清除所有印章", "DeleteAllSeal()", 21);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Excel/AddSeal/save");//设置处理文件保存的请求方法

        //打开Excel文档
        poCtrl.webOpen("/api/doc/InsertSeal/Excel/AddSeal/test5.xls", OpenModeType.xlsNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping(value = "AddSign/Excel1")
    public String showExcel11(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "AddHandSign()", 3);
        poCtrl.addCustomToolButton("删除签字", "DeleteHandSign()", 21);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Excel/AddSign/save");//设置处理文件保存的请求方法

        //打开Excel文档
        poCtrl.webOpen("/api/doc/InsertSeal/Excel/AddSign/test1.xls", OpenModeType.xlsNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping(value = "AddSign/Excel2")
    public String showExcel12(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()", 3);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Excel/AddSign/save");//设置处理文件保存的请求方法

        //打开Excel文档
        poCtrl.webOpen("/api/doc/InsertSeal/Excel/AddSign/test2.xls", OpenModeType.xlsNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSign/Excel3")
    public String showExcel13(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "AddHandSign()", 3);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);

        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Excel/AddSign/save");//设置处理文件保存的请求方法

        //打开Excel文档
        poCtrl.webOpen("/api/doc/InsertSeal/Excel/AddSign/test3.xls", OpenModeType.xlsNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping("AddSeal/save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "InsertSeal/Excel/AddSeal/" + fs.getFileName());
        fs.close();
    }

    @RequestMapping("AddSign/save")
    public void save2(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "InsertSeal/Excel/AddSign/" + fs.getFileName());
        fs.close();
    }


}
