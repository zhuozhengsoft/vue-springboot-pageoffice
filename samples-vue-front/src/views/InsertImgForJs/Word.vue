<template>
	<div class="Word">
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
	      axios.post("/api/InsertImgForJs/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  AfterDocumentOpened() {
			  /**
			   *document.getElementById("PageOfficeCtrl1").InsertWebImage( ImageURL, Transparent, Position );
			   *ImageURL  字符串类型，图片的路径。
			   *Transparent  布尔类型，可选参数，图片是否透明。默认值：FALSE，图片不透明；TRUE表示图片透明。注意：透明色为白色。
			   *Position  整数类型，可选参数，浮于文字上方还是下方。默认值：4，图片浮于文字上方。 5，表示图片衬于文字下方。
			   */
			  //该方法默认插入图片到当前光标处，如果想插入到文档指定位置，可以在文档中插入一个书签来设置位置，然后先定位光标到书签处再插入图片
			  locateBookMark();
			  document.getElementById("PageOfficeCtrl1").InsertWebImage("/api/doc/InsertImgForJs/logo.jpg", false, 4);
		  },
	  
		  //定位书签到光标处
		  locateBookMark() {
			  //获取书签名称
			  var bkName = "PO_logo";
			  //将光标定位到书签所在的位置
			  var mac = "Function myfunc()" + " \r\n"
				  + "  ActiveDocument.Bookmarks(\"" + bkName + "\").Select " + " \r\n"
				  + "End Function " + " \r\n";
			  document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", mac);
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.AfterDocumentOpened = this.AfterDocumentOpened;
			window.locateBookMark = this.locateBookMark;
	    }
	}
</script>

