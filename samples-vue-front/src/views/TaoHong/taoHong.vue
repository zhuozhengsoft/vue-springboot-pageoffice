<template>
	<div class="Word">
		<div id="content">
		    <div id="textcontent" style="width: 1000px; height: 800px;">
		        <div class="flow4">
		            <a href="#" onclick="window.external.close();"> <img alt="返回" src="../../../public/images/return.gif"
		                                                                 border="0"/>文件列表</a>
		            <span style="width: 100px;"> </span><strong>文档主题：</strong>
		            <span style="color: Red;">测试文件</span>
		            <form id="form1">
		                <strong>模板列表：</strong>
		                <span style="color: Red;"> <select name="templateName"
		                                                   id="templateName" style='width: 240px;'>
										<option value='temp2008.doc' selected="selected">
											模板一
										</option>
										<option value='temp2009.doc'>
											模板二
										</option>
										<option value='temp2010.doc'>
											模板三
										</option>
									</select> </span>
		                <span style="color: Red;"><input type="button" value="一键套红"
		                                                 @click="taoHong()"/> </span>
		                <span style="color: Red;"><input type="button" value="保存关闭"
		                                                 @click="saveAndClose()"/> </span>
		            </form>
		        </div>
				
		        <div style="height: 800px; width: auto" v-html="poHtmlCode" />
		
		
		    </div>
		</div>
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'Word',
	    data(){
	      return {
	        poHtmlCode: '',
			form1:''
	      }
	    },
		created: function(){
		  //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
		  axios.post("/api/TaoHong/taoHong").then((response) => {
		    this.poHtmlCode = response.data;
			
		  }).catch(function (err) {
		    console.log(err)
		  })
		},
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      
	              //套红
	              taoHong() {
	                  var mb = document.getElementById("templateName").value;
					   axios.post("/api/TaoHong/taoHong?mb="+mb).then((response) => {
						this.poHtmlCode = response.data;
						
					  }).catch(function (err) {
						console.log(err)
					  })
													
					// document.getElementById("form1").value = this.form1;
	                  // document.forms[0].submit();
	              },
	      
	              //保存并关闭
	              saveAndClose() {
	                  document.getElementById("PageOfficeCtrl1").WebSave();
	                  window.external.close();
	              }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
		  window.taoHong = this.taoHong;
		  window.saveAndClose = this.saveAndClose;
		}
	}
</script>


