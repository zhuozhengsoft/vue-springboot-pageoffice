<template>
<div class="Word">
	
	<p style=" text-indent:10px;">
	    PageOffice 4.0增加了全新的文件打开方式“POBrowser 方式”，此方式提供了更完美的浏览器兼容性解决方案。
	</p>
	<p style=" text-indent:10px;">
	    常规打开文档超链接的代码写法：&lt;a href=&quot;Word?id=12&quot;&gt;某某公司公文-12&lt;/a&gt;</p>
	<p style=" text-indent:10px;">
	    POBrowser打开文档超链接的代码写法：
	    &lt;a href=&quot;<span style=" background-color:#D2E9FF;">javascript:POBrowser.openWindowModeless( &quot;</span>Word?id=12<span
	        style=" background-color:#D2E9FF;">&quot;,&quot;width=800px;height=800px;&quot;)&gt;</span>
	    某某公司公文-12&lt;/a&gt;
	    &nbsp;</p>
	文档标题：<input id="Text1" type="text" size="50"/>
	<input class="button orange" type="button" value="修改" @click="SetTitleText();"/>
	
  <div style="width:auto; height:700px;" v-html="poHtmlCode" >
  </div>
</div>
</template>

<script>
  const axios=require('axios');
  export default{
    name: 'Word',
    data(){
      return {
        message: ' ',
        poHtmlCode: '',

      }
    },
    created: function(){
      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
      axios.post("/api/POBrowser2/Word").then((response) => {
        this.poHtmlCode = response.data;

      }).catch(function (err) {
        console.log(err)
      })
    },
    methods:{
      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      AfterDocumentOpened() {
	          document.getElementById("Text1").value = document.getElementById("PageOfficeCtrl1").DataRegionList.GetDataRegionByName("PO_Title").Value;
	      },
	  
	      SetTitleText() {
	          document.getElementById("PageOfficeCtrl1").DataRegionList.GetDataRegionByName("PO_Title").Value = document.getElementById("Text1").value;
	      },
	  
	      Close() {
	          window.external.close();
	      },
		  
		  Save() {
			  document.getElementById("PageOfficeCtrl1").WebSave();
		  },
	  
	      //文档关闭前先提示用户是否保存
	      BeforeBrowserClosed() {
	          if (document.getElementById("PageOfficeCtrl1").IsDirty) {
	              if (confirm("提示：文档已被修改，是否继续关闭放弃保存 ？")) {
	                  return true;
	  
	              } else {
	  
	                  return false;
	              }
	  
	          }
	  
	      }
    },
    mounted: function(){
      // 将vue中的方法赋值给window
	  window.Close = this.Close;
	  window.Save = this.Save;
      window.AfterDocumentOpened = this.AfterDocumentOpened;
	  window.BeforeBrowserClosed = this.BeforeBrowserClosed;
    }
}
</script>
<style type="text/css">
        .button {
            display: inline-block;
            zoom: 1; /* zoom and *display = ie7 hack for display:inline-block */
            *display: inline;
            vertical-align: baseline;
            margin: 0 2px;
            outline: none;
            cursor: pointer;
            text-align: center;
            text-decoration: none;
            font: 14px/100% Arial, Helvetica, sans-serif;
            padding: .25em 1em .25em;
            text-shadow: 0 1px 1px rgba(0, 0, 0, .3);
            -webkit-border-radius: .5em;
            -moz-border-radius: .5em;
            border-radius: .5em;
            -webkit-box-shadow: 0 1px 2px rgba(0, 0, 0, .2);
            -moz-box-shadow: 0 1px 2px rgba(0, 0, 0, .2);
            box-shadow: 0 1px 2px rgba(0, 0, 0, .2);
        }

        .button:hover {
            text-decoration: none;
        }

        .button:active {
            position: relative;
            top: 1px;
        }

        .bigrounded {
            -webkit-border-radius: 2em;
            -moz-border-radius: 2em;
            border-radius: 2em;
        }

        .medium {
            font-size: 12px;
            padding: .4em 1.5em .42em;
        }

        .small {
            font-size: 11px;
            padding: .2em 1em .275em;
        }

        /* orange */
        .orange {
            color: #fef4e9;
            border: solid 1px #da7c0c;
            background: #f78d1d;
            background: -webkit-gradient(linear, left top, left bottom, from(#faa51a), to(#f47a20));
            background: -moz-linear-gradient(top, #faa51a, #f47a20);
            filter: progid:DXImageTransform.Microsoft.gradient(startColorstr='#faa51a', endColorstr='#f47a20');
        }

        .orange:hover {
            background: #f47c20;
            background: -webkit-gradient(linear, left top, left bottom, from(#f88e11), to(#f06015));
            background: -moz-linear-gradient(top, #f88e11, #f06015);
            filter: progid:DXImageTransform.Microsoft.gradient(startColorstr='#f88e11', endColorstr='#f06015');
        }

        .orange:active {
            color: #fcd3a5;
            background: -webkit-gradient(linear, left top, left bottom, from(#f47a20), to(#faa51a));
            background: -moz-linear-gradient(top, #f47a20, #faa51a);
            filter: progid:DXImageTransform.Microsoft.gradient(startColorstr='#f47a20', endColorstr='#faa51a');
        }
    </style>