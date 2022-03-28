<template>
	<div class="Excel">
		<div style="font-size:12px; line-height:20px; border-bottom:dotted 1px #ccc;border-top:dotted 1px #ccc; padding:5px;">
		    <span style="color:red;">操作说明：</span>请点击按钮。
		    <input id="Button1" type="button" value="获取Sheet1中B4单元格的值" @click="Button1_onclick()"/>
		    <input id="Button2" type="button" value="设置Sheet1中C4单元格的值为：100" @click="Button2_onclick()"/>
		    <br/>
		
		    关键代码：看js函数<span style="background-color:Yellow;">getCellValue(sheet, cell)&nbsp;&nbsp; setCellValue(sheet, cell, value)</span>
		</div>
		<br/>
		
	  <div style="height: 800px; width: auto" v-html="poHtmlCode" />
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'Excel',
	    data(){
	      return {
	        poHtmlCode: '',
	
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/JsOpXlsCellText/Excel").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  setCellValue(sheet, cell, value) {
			  var sMac = "function myfunc()" + "\r\n"
				  + "Application.Sheets(\"" + sheet + "\").Range(\"" + cell + "\").Value = \"" + value + "\" \r\n"
				  + "End function";
			  return document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", sMac);
		  },
	  
		  getCellValue(sheet, cell) {
			  var sMac = "function myfunc()" + "\r\n"
				  + "myfunc = Application.Sheets(\"" + sheet + "\").Range(\"" + cell + "\").Text \r\n"
				  + "End function";
			  return document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", sMac);
		  },
	  
		  Button1_onclick() {
			  document.getElementById("PageOfficeCtrl1").Alert(getCellValue("Sheet1", "B4"));
		  },
	  
		  Button2_onclick() {
			  setCellValue("Sheet1", "C4", "100");
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.setCellValue = this.setCellValue;
			window.getCellValue = this.getCellValue;
			
	    }
	}
</script>

