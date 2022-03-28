package com.zhuozhengsoft.samplesspringbootback.controller.insertSeal;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.samplesspringbootback.util.GetDirPathUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@RestController
@RequestMapping(value = "/InsertSeal/Word/")
public class WordInsertSealController {
    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";


    @RequestMapping(value = "BatchAddSeal/index")
    public String showindex(HttpServletRequest request) {
        return dir + "InsertSeal/Word/BatchAddSeal/";
    }

    @RequestMapping(value = "BatchAddSeal/Edit")
    public String showBatchAddSealEdit(HttpServletRequest request) {
        String filePath = "";
        String id = request.getParameter("id").trim();
        if ("1".equals(id)) {
            filePath = "test1.doc";
        }
        if ("2".equals(id)) {
            filePath = "test2.doc";
        }
        if ("3".equals(id)) {
            filePath = "test3.doc";
        }
        if ("4".equals(id)) {
            filePath = "test4.doc";
        }

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/BatchAddSeal/save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/BatchAddSeal/" + filePath, OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "BatchAddSeal/AddSeal")
    public String showBatchAddSealAddSeal(HttpServletRequest request) {

        String path = request.getContextPath();
        String filePath = "";
        String id = request.getParameter("id").trim();
        if ("1".equals(id)) {
            filePath = "test1.doc";
        }
        if ("2".equals(id)) {
            filePath = "test2.doc";
        }
        if ("3".equals(id)) {
            filePath = "test3.doc";
        }
        if ("4".equals(id)) {
            filePath = "test4.doc";
        }

        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage("/api/poserver.zz");
        fmCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        //设置保存页面
        fmCtrl.setSaveFilePage("/api/InsertSeal/Word/BatchAddSeal/save");//设置处理文件保存的请求方法
        fmCtrl.fillDocument("/api/doc/InsertSeal/Word/BatchAddSeal/" + filePath, DocumentOpenType.Word);
        return fmCtrl.getHtmlCode("FileMakerCtrl1");
    }


    @RequestMapping("BatchAddSeal/save")
    public void save3(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "InsertSeal/Word/BatchAddSeal/" + fs.getFileName());
        fs.close();
    }


    @RequestMapping(value = "AddSeal/Word1")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

//添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSeal/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSeal/test1.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping(value = "AddSeal/Word2")
    public String showWord2(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        //设置保存页面
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSeal/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSeal/test2.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSeal/Word3")
    public String showWord3(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);

        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSeal/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSeal/test3.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSeal/Word4")
    public String showWord4(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("添加印章位置", "InsertSealPos()", 2);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSeal/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSeal/test4.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSeal/Word5")
    public String showWord5(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("盖章到印章位置", "AddSealByPos()", 2);

        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSeal/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSeal/test5.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSeal/Word6")
    public String showWord6(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("盖章到印章位置", "AddSealByPos()", 2);

        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSeal/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSeal/test6.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSeal/Word7")
    public String showWord7(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage( "/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("盖章到印章位置", "AddSealByPos()", 2);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSeal/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSeal/test7.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSeal/Word8")
    public String showWord8(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSeal/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSeal/test8.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSeal/Word9")
    public String showWord9(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);

        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSeal/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSeal/test9.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSeal/Word10")
    public String showWord10(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage( "/api/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除指定印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("清除所有印章", "DeleteAllSeal()", 21);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSeal/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSeal/test10.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSign/Word1")
    public String showWord11(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

//添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()", 3);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSign/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSign/test1.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping(value = "AddSign/Word2")
    public String showWord12(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()", 2);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSign/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSign/test2.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSign/Word3")
    public String showWord13(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()", 3);

        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSign/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSign/test3.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSign/Word4")
    public String showWord14(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "AddHandSign()", 3);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSign/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSign/test4.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "AddSign/Word5")
    public String showWord15(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
//添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()", 3);

        //设置保存页面
        poCtrl.setSaveFilePage("/api/InsertSeal/Word/AddSign/save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/doc/InsertSeal/Word/AddSign/test5.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping("AddSeal/save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "InsertSeal/Word/AddSeal/" + fs.getFileName());
        fs.close();
    }

    @RequestMapping("AddSign/save")
    public void save2(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "InsertSeal/Word/AddSign/" + fs.getFileName());
        fs.close();
    }

}
