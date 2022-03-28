package com.zhuozhengsoft.samplesspringbootback;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.web.servlet.ServletRegistrationBean;
import org.springframework.context.annotation.Bean;

import java.io.FileNotFoundException;


@SpringBootApplication
public class SamplesSpringbootBackApplication {

    @Value("${posyspath}")
    private String poSysPath;

    public static void main(String[] args) {
        SpringApplication.run(SamplesSpringbootBackApplication.class, args);

    }
    /**
     * 添加PageOffice的服务器端授权程序Servlet（必须）
     *
     * @return
     */
    @Bean
    public ServletRegistrationBean pageofficeRegistrationBean() {
        com.zhuozhengsoft.pageoffice.poserver.Server poserver = new com.zhuozhengsoft.pageoffice.poserver.Server();
        poserver.setSysPath(poSysPath);//设置PageOffice注册成功后,license.lic文件存放的目录
        ServletRegistrationBean srb = new ServletRegistrationBean(poserver);
        srb.addUrlMappings("/poserver.zz");
        srb.addUrlMappings("/posetup.exe");
        srb.addUrlMappings("/pageoffice.js");
        srb.addUrlMappings("/jquery.min.js");
        srb.addUrlMappings("/pobstyle.css");
        srb.addUrlMappings("/sealsetup.exe");
        return srb;//
    }


    /**
     * 添加印章管理程序Servlet（可选）
     *
     * @return
     */
    @Bean
    public ServletRegistrationBean zoomsealRegistrationBean() throws FileNotFoundException {
        com.zhuozhengsoft.pageoffice.poserver.AdminSeal adminSeal = new com.zhuozhengsoft.pageoffice.poserver.AdminSeal();
        adminSeal.setAdminPassword(poSysPath);//设置印章管理员admin的登录密码（为了安全起见，强烈建议修改此密码）
        /**如果当前项目是打成jar或者war包运行，强烈建议将poseal.db文件的路径更换成某个固定的绝对路径下,不要放当前项目文件夹下,为了防止每次重新打包程序导致poseal.db被替换的问题。
         * 比如windows服务器下：D:/lic/，linux服务器下:/root/lic/
         */
        //设置印章数据库文件poseal.db存放的目录
        adminSeal.setSysPath(poSysPath);
        ServletRegistrationBean srb = new ServletRegistrationBean(adminSeal);
        srb.addUrlMappings("/adminseal.zz");
        srb.addUrlMappings("/sealimage.zz");
        srb.addUrlMappings("/loginseal.zz");
        return srb;//
    }
}
