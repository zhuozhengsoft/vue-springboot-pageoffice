<template>
	<div class="Word">
		<div style="font-size:12px; line-height:20px; border-bottom:dotted 1px #ccc;border-top:dotted 1px #ccc; padding:5px;">
		    <span style="color:red;">操作说明：</span>定位word文件中的光标到想插入超链接的位置，然后点“插入超链”按钮。<br/>
		
		    关键代码：看js函数<span style="background-color:Yellow;">addHyperLink()</span>
		</div><br/>
		
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
	      axios.post("/api/WordHyperLink/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
			//  addHyperLink()
			//  {
			// 	 var docObj = document.getElementById("PageOfficeCtrl1").Document;
			// 	 docObj.Application.ActiveWindow.View.ShowFieldCodes = false; //不要以域代码的形式显示超链接
			// 	docObj.Hyperlinks.Add(docObj.Application.Selection.Range, "http://www.zhuozhengsoft.com/", "", "", "卓正");
			// }
	      
		  addHyperLink() {
			  var text = "卓正志远";
			  var url = "http://www.zhuozhengsoft.com/";
	  
			  var mac = "Function myfunc()" + " \r\n"
				  + "  Application.ActiveWindow.View.ShowFieldCodes = False " + " \r\n"
				  + "  ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _" + " \r\n"
				  + "   \"" + url + "\", SubAddress:=\"\", ScreenTip:=\"\", _ " + " \r\n"
				  + "   TextToDisplay:=\"" + text + "\" " + " \r\n"
				  + "End Function " + " \r\n";
			  document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", mac);
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.addHyperLink = this.addHyperLink;
	    }
	}
</script>

