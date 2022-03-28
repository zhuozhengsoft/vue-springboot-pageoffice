package com.zhuozhengsoft.samplesspringbootback.controller;

import com.zhuozhengsoft.pageoffice.DocumentVersion;
import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.samplesspringbootback.util.Doc;
import com.zhuozhengsoft.samplesspringbootback.util.GetDirPathUtil;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.sql.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Date;

@RestController
@RequestMapping(value = "/CreateWord/")
public class CreateWordController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "index")
    public Map showWord20(HttpServletRequest request) throws ClassNotFoundException, SQLException, ParseException, FileNotFoundException, UnsupportedEncodingException {

     Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/CreateWord.db";

        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("select * from word order by id desc");
        String fileName = "";
        String subject = "";
        String submitTime = "";
        List<Doc> list = new ArrayList<Doc>();

        while (rs.next()) {
            int id = rs.getInt("ID");
            fileName = rs.getString("FileName");
            subject = rs.getString("Subject");
            submitTime = rs.getString("SubmitTime");
            if (submitTime != null && submitTime.length() > 0) {
                submitTime = new SimpleDateFormat("yyyy/MM/dd")
                        .format(new SimpleDateFormat("yyyy-MM-dd")
                                .parse(submitTime));
            }
            Doc doc = new Doc();
            doc.setId(id);
            doc.setFileName(fileName);
            doc.setSubject(subject);
            doc.setSubmitTime(submitTime);
            list.add(doc);

        }
        rs.close();
        stmt.close();
        conn.close();
        HashMap<Object, Object> map = new HashMap<>();
        map.put("list", list);
        //--- PageOffice的调用代码 结束 -----
        return map;
    }

    @RequestMapping(value = "Word")
    public Map showWord(HttpServletRequest request) throws ClassNotFoundException, FileNotFoundException, SQLException {
        String subject = "";
        String fileName = "";
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/CreateWord.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        String id = request.getParameter("id");
        if (!id.equals("") && !id.equals(null)) {
            ResultSet rs = stmt.executeQuery("select * from word where ID=" + id);
            subject = rs.getString("Subject");
            fileName = rs.getString("FileName");
            rs.close();
        }
        stmt.close();
        conn.close();


        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("/api/CreateWord/save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/api/doc/CreateWord/" + fileName, OpenModeType.docNormalEdit, "张三");
        HashMap<Object, Object> map = new HashMap<>();
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("subject", subject);
        return map;
    }


    @RequestMapping(value = "creatWord")
    public String showWord3(HttpServletRequest request) throws ClassNotFoundException, FileNotFoundException, SQLException {


        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/api/poserver.zz");//设置服务页面
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);
        poCtrl.setJsFunction_BeforeDocumentSaved("BeforeDocumentSaved()");

        //设置保存页面
        poCtrl.setSaveFilePage("/api/CreateWord/SaveNewFile");//设置处理文件保存的请求方法

        //新建Word文件，webCreateNew方法中的两个参数分别指代“操作人”和“新建Word文档的版本号”
        poCtrl.webCreateNew("张佚名", DocumentVersion.Word2003);

        return poCtrl.getHtmlCode("PageOfficeCtrl1");
    }


    @RequestMapping("SaveNewFile")
    public void save2(HttpServletRequest request, HttpServletResponse response) throws ClassNotFoundException, FileNotFoundException, SQLException {
        FileSaver fs = new FileSaver(request, response);


        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/CreateWord.db";

        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("select Max(ID) from word");
        int newID = 1;
        if (rs.next()) {
            newID = Integer.parseInt(rs.getString(1)) + 1;
        }
        rs.close();

        String FileSubject = fs.getFormField("FileSubject").trim();
        String fileName = "aabb" + newID + ".doc";
        String getFile = (String) request.getParameter("FileSubject");
        //if (getFile != null && getFile.length() > 0) FileSubject = new String(getFile.getBytes("iso-8859-1"));
        //out.print(FileSubject);
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//设置日期格式
        // new Date()为获取当前系统时间
        String strsql = "Insert into word(ID,FileName,Subject,SubmitTime) values("
                + newID
                + ",'"
                + fileName
                + "','"
                + FileSubject
                + "','"
                + df.format(new Date()) + "')";
        stmt.executeUpdate(strsql);
        stmt.close();
        conn.close();

        fs.saveToFile(dir + "CreateWord/" + fileName);
        //设置保存结果
        fs.setCustomSaveResult("ok");
        fs.close();
    }

    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "CreateWord/" + fs.getFileName());
        fs.close();
    }

}
