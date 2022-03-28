package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegionInsertType;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import com.zhuozhengsoft.samplesspringbootback.util.GetDirPathUtil;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.*;
import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping(value = "/ExaminationPaper/")
public class ExaminationPaperController {
    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "index")
    public StringBuilder showindex(HttpServletRequest request, Map<String, Object> map) throws SQLException, FileNotFoundException, ClassNotFoundException {
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/ExaminationPaper.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("Select * from stream");
        boolean flg = false;//标识是否有数据
        StringBuilder strHtmls = new StringBuilder();
        strHtmls.append("<tr  style='background-color:#55aa7f;'>");
        strHtmls.append("<td style='text-align:center;width=10%' >选择</td>");
        strHtmls.append("<td style='text-align:center;width=30%'>题库编号</td>");
        strHtmls.append("<td style='text-align:center;width=60%'>操作</td>");
        strHtmls.append("</tr>");
        while (rs.next()) {
            flg = true;
            String pID = rs.getString("ID");
            strHtmls.append("<tr  style='background-color:#f5ffe1;'>");
            strHtmls.append("<td style='text-align:center'><input id='check" + pID + "'  type='checkbox' /></td>");
            strHtmls.append("<td style='text-align:center'>选择题-" + pID + "</td>");
            strHtmls.append("<td style='text-align:center'><a href='javascript:POBrowser.openWindowModeless(\"Word?id=" + pID + "\" ,\"width=1200px;height=800px;\");'>编辑</a></td>");
            strHtmls.append("</tr>");
        }

        if (!flg) {
            strHtmls.append("<tr>\r\n");
            strHtmls.append("<td width='100%' height='100' align='center'>对不起，暂时没有可以操作的数据。\r\n");
            strHtmls.append("</td></tr>\r\n");
        }

        return strHtmls;
    }


    @RequestMapping(value = "Word")
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

         //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //设置保存页面
        String id = request.getParameter("id");

        //设置保存页面
        poCtrl.setSaveFilePage("/api/ExaminationPaper/save?id=" + id);//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/api/ExaminationPaper/Openfile?id=" + id, OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping(value = "Compose")
    public Map showCompose(HttpServletRequest request) {
        String idlist = request.getParameter("ids").trim();
        String[] ids = idlist.split(",");//将idlist按照","截取后存到ids数组中，然后遍历数组用js插入文件即可

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面

        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.setCustomToolbar(false);
        poCtrl.setCaption("生成试卷");
        poCtrl.setJsFunction_AfterDocumentOpened("Create()");

        //打开Word文档
        poCtrl.webOpen("/api/doc/ExaminationPaper/test.doc", OpenModeType.docNormalEdit, "张三");
        Map<Object, Object> map = new HashMap<>();
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("idArr", ids);
        return map;
    }


    @RequestMapping(value = "Compose2")
    public String showCompose2(HttpServletRequest request) {
        String idlist = request.getParameter("ids").trim();
        String[] ids = idlist.split(","); //将idlist按照","截取后存到ids数组中，然后遍历数组用js插入文件
        String temp = "PO_begin";//存储数据区域名称
        int num = 1;//试题编号
        WordDocument doc = new WordDocument();
        for (int i = 0; i < ids.length; i++) {

            DataRegion dataNum = doc.createDataRegion("PO_" + num,
                    DataRegionInsertType.After, temp);
            dataNum.setValue(num + ".\t");
            DataRegion dataRegion = doc.createDataRegion("PO_begin"
                    + (i + 1), DataRegionInsertType.After, "PO_" + num);
            dataRegion.setValue("[word]/api/ExaminationPaper/Openfile?id=" + ids[i]
                    + "[/word]");
            temp = "PO_begin" + (i + 1);
            num++;
        }

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.setCustomToolbar(false);
        poCtrl.setCaption("生成试卷");
        poCtrl.setWriter(doc);

        //打开Word文档
        poCtrl.webOpen("/api/doc/ExaminationPaper/test.doc", OpenModeType.docNormalEdit, "张三");
        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }

    @RequestMapping(value = "Openfile")
    public void openWord(HttpServletRequest request, HttpServletResponse response) throws SQLException, ClassNotFoundException, IOException {
        String err = "";
        if (request.getParameter("id") != null
                && request.getParameter("id").trim().length() > 0) {
            String id = request.getParameter("id");
            Class.forName("org.sqlite.JDBC");
            String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/ExaminationPaper.db";
            Connection conn = DriverManager.getConnection(strUrl);
            Statement stmt = conn.createStatement();
            String strSql = "select * from stream where id =" + id;
            ResultSet rs = stmt.executeQuery(strSql);
            if (rs.next()) {
                //******读取磁盘文件，输出文件流 开始*******************************
                byte[] imageBytes = rs.getBytes("Word");
                int fileSize = imageBytes.length;

                response.reset();
                response.setContentType("application/msword"); // application/x-excel, application/ms-powerpoint, application/pdf
                response.setHeader("Content-Disposition", "attachment; filename=down.doc"); //fileN应该是编码后的(utf-8)
                response.setContentLength(fileSize);
                OutputStream outputStream = response.getOutputStream();
                outputStream.write(imageBytes);
                outputStream.flush();
                outputStream.close();
                outputStream = null;

            } else {
                err = "未获得文件的信息";
            }
            rs.close();
            stmt.close();
            conn.close();
        } else {
            err = "未获得文件的ID";
            //out.print(err);
        }
        if (err.length() > 0)
            err = "<script>alert(" + err + ");</script>";

    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "ExaminationPaper/" + fs.getFileName());
        fs.close();
    }
}
