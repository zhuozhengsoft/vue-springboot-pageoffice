<template>
	<div class="Word">
		<div style=" font-size:small; color:Red;">
		    <label>关键代码：看js函数：</label>
		    <label>function locateBookMark()</label>
		    <br/>
		    <label>将光标定位到书签前，请先在文本框中输入要定位到的书签名称（可点击Office工具栏上的“插入”→“书签”，来查看当前Word文档中所有的书签名称）！</label><br/>
		    <label>书签名称：</label><input id="txtBkName" type="text" value="PO_seal"/>
		</div>
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
	      axios.post("/api/WordLocateBKMK/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
			 // locateBookMark() {
				//  //获取书签名称
				//  var bkName = document.getElementById("txtBkName").value;
				//  //将光标定位到书签所在的位置
				//  document.getElementById("PageOfficeCtrl1").Document.Bookmarks(bkName).Select();
			 // }
			  locateBookMark() {
				  //获取书签名称
				  var bkName = document.getElementById("txtBkName").value;
				  //将光标定位到书签所在的位置
				  var mac = "Function myfunc()" + " \r\n"
					  + "  ActiveDocument.Bookmarks(\"" + bkName + "\").Select " + " \r\n"
					  + "End Function " + " \r\n";
				  document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", mac);
			  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.locateBookMark = this.locateBookMark;
	    }
	}
</script>

