<template>
	<div class="Word">
		    function AfterDocumentOpened() {
		        document.getElementById("Text1").value = document.getElementById("PageOfficeCtrl1").DataRegionList.GetDataRegionByName("PO_Title").Value;
		    }
		
		    function setTitleText() {
		        document.getElementById("PageOfficeCtrl1").DataRegionList.GetDataRegionByName("PO_Title").Value = document.getElementById("Text1").value;
		    }
		</script>
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
		<input type="button" value="修改" onclick="setTitleText();"/>
	  <div style="height: 800px; width: auto" v-html="poHtmlCode" />
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'Word',
	    data(){
	      return {
	        poHtmlCode: '',
	
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/POBrowserTopic/Word1").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
				Save() {
					document.getElementById("PageOfficeCtrl1").WebSave();
			
				},
				AfterDocumentOpened() {
					document.getElementById("Text1").value = document.getElementById("PageOfficeCtrl1").DataRegionList.GetDataRegionByName("PO_Title").Value;
				},
			
				setTitleText() {
					document.getElementById("PageOfficeCtrl1").DataRegionList.GetDataRegionByName("PO_Title").Value = document.getElementById("Text1").value;
				}
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.Save = this.Save;
			window.AfterDocumentOpened = this.AfterDocumentOpened;
			window.setTitleText = this.setTitleText;
		
	    }
	}
</script>

