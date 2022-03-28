<template>
	<div class="Word">
		    <div style=" font-size:small; color:Red;">
		        <label>关键代码：看js函数：</label>
		        <label>function addBookMark() 和 function delBookMark()</label>
		        <br/>
		        <label>插入书签时，请先输入要插入的书签名称和文本；删除书签时，请先输入相应的书签名称！</label><br/>
		        <label>书签名称：</label><input id="txtBkName" type="text" value="test"/>
		        &nbsp;&nbsp;<label>书签文本：</label><input id="txtBkText" type="text" value="[测试]"/>
		    </div>
		    <input id="Button1" type="button" @click="addBookMark();" value="插入书签"/>
		    <input id="Button2" type="button" @click="delBookMark()" value="删除书签"/>
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
	      axios.post("/api/WordAddBKMK/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  addBookMark() {
			  var bkName = document.getElementById("txtBkName").value;
			  var bkText = document.getElementById("txtBkText").value;
	  
			  var mac = "Function myfunc()" + " \r\n"
				  + "Dim r As Range " + " \r\n"
				  + "Set r = Application.Selection.Range " + " \r\n"
				  + "r.Text = \"" + bkText + "\"" + " \r\n"
				  + "Application.ActiveDocument.Bookmarks.Add Name:=\"" + bkName + "\", Range:=r " + " \r\n"
				  + "Application.ActiveDocument.Bookmarks(\"" + bkName + "\").Select " + " \r\n"
				  + "End Function " + " \r\n";
			  document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", mac);
		  },
	  
		  delBookMark() {
			  var bkName = document.getElementById("txtBkName").value;
			  var mac = "Function myfunc()" + " \r\n"
				  + "If  Application.ActiveDocument.Bookmarks.Exists(\"" + bkName + "\") Then " + " \r\n"
				  + "    Application.ActiveDocument.Bookmarks(\"" + bkName + "\").Select " + " \r\n"
				  + "    Application.Selection.Range.Text = \"\" " + " \r\n"
				  + "End If " + " \r\n"
				  + "End Function " + " \r\n";
			  document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", mac);
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.addBookMark = this.addBookMark;
			window.delBookMark = this.delBookMark;
	    }
	}
</script>

