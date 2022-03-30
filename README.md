# vue-springboot-pageoffice
### 一、简介

​       vue-springboot-pageoffice项目演示了在前端Vue框架和后端Springboot框架的结合下如何使用PageOffice产品，此项目是一个demo项目。

### 二、项目环境要求

- 前端Vue项目（samples-vue-front）：Node.js10及以上版本
- 后端Springboot项目（samples-springboot-back）：Intelij IDEA,jdk1.8及以上版本

### 三、项目运行准备

   在当前服务器磁盘上新建一个pageoffice系统文件夹，例如：D:/pageoffice（此文件夹将用来存放PageOffice注册后生成的授权文件“license.lic”）。

### 四、项目运行步骤

- ##### 前端Vue项目（samples-vue-front）

1. **npm install** ：安装依赖
2. **npm run dev** ：运行启动

- ##### 后端Springboot项目（samples-springboot-back）

1. 使用git clone或者直接下载项目压缩包到本地并解压缩。
2. 打开application.properties文件，将posyspath变量的值配置成您上一步新建的PageOffice系统文件夹  （例如：D:/pageoffice）。
3. 如果您要测试PageOffice的电子印章功能，请拷贝当前项目根目录下的poseal.db文件到PageOffice系统文件夹下（例如：D:/pageoffice/poseal.db）。
4. 运行项目：点击运行按钮即可。

> 注意：如果后端Springboot项目的8081端口已经被占用，后端项目更换其他端口后，请记得将前端Vue项目samples-vue-front/config下中的index.js中的 proxyTable对象中的 target指向的地址改成更改后的后端项目的端口，并且将前端Vue项目的index.html中对pageoffice.js引用时的后端项目的端口也改成更改后的端口。

### 五、PageOffice序列号

​     PageOfficeV5.0标准版试用序列号：I2BFU-MQ89-M4ZZ-ZWY7K           
​     PageOfficeV5.0专业版试用序列号：DJMTF-HYK4-BDQ3-2MBUC

### 六、集成PageOffice到您的项目中的关键步骤

- ##### 后端Springboot项目

1. 在您项目的pom.xml中通过下面的代码引入PageOffice依赖。

> pageoffice.jar的所有Releases版本都已上传到了Maven中央仓库，具体您要引用哪个版本请在Maven中央仓库地址中查看，建议使用Maven中央仓库中pageoffice.jar的最新版本。(Maven中央仓库中pageoffice.jar的地址：https://mvnrepository.com/artifact/com.zhuozhengsoft/pageoffice)

```
<dependency>
     <groupId>com.zhuozhengsoft</groupId>   
  <artifactId>pageoffice</artifactId>   
  <version>5.3.0.4</version>
</dependency>
```

2. 在您项目的启动类Application类中配置如下代码。

```
@Bean
   public ServletRegistrationBean pageofficeRegistrationBean()  {
com.zhuozhengsoft.pageoffice.poserver.Server poserver = new com.zhuozhengsoft.pageoffice.poserver.Server();
/**如果当前项目是打成jar或者war包运行，强烈建议将license的路径更换成某个固定的绝
*对路径下，不要放当前项目文件夹下,为了防止每次重新发布项目导致license丢失问题。
* 比如windows服务器下：D:/pageoffice，linux服务器下:/root/pageoffice
 */
 //设置PageOffice注册成功后,license.lic文件存放的目录
 poserver.setSysPath(poSysPath);
 //poSysPath可以在application.properties这个文件中配置，也可以直设置文件夹路径，比如：poserver.setSysPath("D:/pageoffice");
  ServletRegistrationBean srb = new ServletRegistrationBean(poserver);
  srb.addUrlMappings("/poserver.zz");
  srb.addUrlMappings("/posetup.exe");
  srb.addUrlMappings("/pageoffice.js");
  srb.addUrlMappings("/jquery.min.js");
  srb.addUrlMappings("/pobstyle.css");
  srb.addUrlMappings("/sealsetup.exe");
  return srb;
  }
```

3. 新建Controller并调用PageOffice。例如：

```
 @RequestMapping(value="/Word")
 @ResponseBody
    public String showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务页面，/api是前端vue项目的代理地址
        poCtrl.setServerPage("/api/poserver.zz");
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("另存为", "SaveAs", 12);
        poCtrl.addCustomToolButton("打印设置", "PrintSet", 0);
        poCtrl.addCustomToolButton("打印", "PrintFile", 6);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
        poCtrl.addCustomToolButton("-", "", 0);
        poCtrl.addCustomToolButton("关闭", "Close", 21);
        //设置保存方法的url
        poCtrl.setSaveFilePage("/api/SimpleWord/save");
        //打开word
        poCtrl.webOpen("/api/doc/SimpleWord/test.doc", OpenModeType.docNormalEdit, "张三");
        return  poCtrl.getHtmlCode("PageOfficeCtrl1");
    }
```
- ##### 前端Vue项目

1.  在您Vue项目的根目录下index.html中引用后端项目根目录下pageoffice.js文件。例如：

<!--注意：8081是后端项目端口号，samples-springboot-back是项目名称，这些都不是固定死的，根据您后端项目的具体地址具体引用即可。如何判断当前pageoffice.js的引用地址是否正确呢？方法是可以将这个引用pageoffice.js的url地址直接粘贴到浏览器地址栏，如果提示能正确下载到这个js文件，则说明引用地址正确。-->

```
<script type="text/javascript" src="http://localhost:8081/samples-springboot-back/pageoffice.js"></script>
```

2. 在您要打开文件的Vue页面，通过超链接点击或者按钮点击触发调用POBrowser打开一个新的Vue页面。比如通过超链接打开了一个新的Word.vue的空白页面，代码如下：

```
<a href="javascript:POBrowser.openWindowModeless('SimpleWord/Word','width=1150px;height=900px;');">
```

3. 在Word.vue页面中通过vue的create钩子函数通过axios去调用后端打开文件的controller。

```vue
created: function () {
    //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里建议用axios方式去请求后台打开文件的controller方法
    axios
      .post("/api/SimpleWord/Word")
      .then((response) => {
        this.poHtmlCode = response.data;
      })
      .catch(function (err) {
        console.log(err);
      });
  },
```
4. 在Word.vue页面中通过v-html标签将上一步axios请求返回的data输出到当前页面的某个div（这个div用来存放PageOffice控件）中，例如：

```
  <div style="height: 800px; width: auto" v-html="poHtmlCode" />
```

5. 在Word.vue页面的methods方法中添加PageOffice中常用的保存，打印等方法。例如：

```vue
methods: {
//控件中的一些常用方法都在这里调用，比如保存，打印等等
/**
 * Save()，callParent()等方法都是/api/SimpleWord/Word这个后台controller中PageOfficeCtrl控件通过poCtrl.addCustomToolButton定义的方法。
 */
Save() {
  document.getElementById("PageOfficeCtrl1").WebSave();
}
},
```

6. 在Word.Vue页面的mounted方法中将上一步的PageOffice中的自定义保存挂载到window对象上。比如：

```vue
mounted: function () {
// 将PageOffice控件中的方法通过mounted挂载到window对象上，只有挂载后才能被vue组件识别
window.Save = this.Save;
},
```
### 七、电子印章功能说明

​     如果您的项目要用到PageOffice自带电子印章功能，请按下面的步骤进行操作。

​     1.请在您的web项目的启动类Application类中加上印章功能相关配置代码，具体代码请参考当前项目springboot-pageoffice启动类中电子印章功能配置代码，此处不再赘述。

> ​    注意：adminSeal.setSysPath(poSysPath)中的poSysPath指向的地址必须是您当前项目license.lic文件所在的目录。

​    2.请拷贝当前项目根目录下poseal.db文件到您的web项目的license.lic文件所在的文件夹下。

​    3.请参考当前项目的[一、15、演示加盖印章和签字功能（以Word为例）](http://localhost:8080/InsertSeal/index)  示例代码进行盖章功能调用。

### 八、 PageOffice开发帮助

​     1.[Java API文档](http://www.zhuozhengsoft.com/help/java3/index.html) 

​     2.[JS API文档](http://www.zhuozhengsoft.com/help/js3/index.html)  

​     3.[PageOffice从入门到精通](https://www.kancloud.cn/pageoffice_course_group/pageoffice_course/646953)

​     技术支持：http://www.zhuozhengsoft.com/Technical/

### 九、联系我们

​   卓正官网：[http://www.zhuozhengsoft.com](http://www.zhuozhengsoft.com)

​   联系电话：400-6600-770  

   QQ: 800038353