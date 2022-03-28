<template>
	<div class="Word">
		<div style="font-size:12px; line-height:20px; border-bottom:dotted 1px #ccc;border-top:dotted 1px #ccc; padding:5px;">
		    <span style="color:red;">操作说明：</span>请输入页码后点跳转按钮。页码：
		    <input id="pageNum" type="text" value="3"/>
		    <input id="Button1" type="button" value="跳转" @click="Button1_onclick()"/>
		    <br/>
		
		    关键代码：看js函数<span style="background-color:Yellow;">gotoPage(num)</span>
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
	      axios.post("/api/WordGoToPage/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  LTrim(str) {
			  var i;
			  for (i = 0; i < str.length; i++) {
				  if (str.charAt(i) != " " && str.charAt(i) != " ") break;
			  }
			  str = str.substring(i, str.length);
			  return str;
		  },
	  
		  RTrim(str) {
			  var i;
			  for (i = str.length - 1; i >= 0; i--) {
				  if (str.charAt(i) != " " && str.charAt(i) != " ") break;
			  }
			  str = str.substring(0, i + 1);
			  return str;
		  },
	  
		  Trim(str) {
			  return LTrim(RTrim(str));
		  },
		  
		  gotoPage(num) {
			  var sMac = "function myfunc()" + "\r\n"
				  + "  If " + num + " > Application.Selection.Information(4) Then " + "\r\n"
				  + "     Msgbox \"超出文档范围，本文共\" & Application.Selection.Information(4) & \"页\"" + "\r\n"
				  + "  End If " + "\r\n"
				  + "  Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Name:= " + num + "\r\n"
				  + "End function";
			  return document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", sMac);
		  },
		  Button1_onclick() {
			  var num = document.getElementById("pageNum").value;
			  if ("" == Trim(num)) {
				  document.getElementById("PageOfficeCtrl1").Alert("请输入页码");
				  return;
			  }
			  if (isNaN(num)) {
				  document.getElementById("PageOfficeCtrl1").Alert("只能输入数字");
				  return;
			  }
			  gotoPage(num);
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.LTrim = this.LTrim;
			window.RTrim = this.RTrim;
			window.Trim = this.Trim;
			window.gotoPage = this.gotoPage;
	    }
	}
</script>

