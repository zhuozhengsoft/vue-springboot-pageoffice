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
	      axios.post("/api/InsertImgWaterMark/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  AfterDocumentOpened() {
			  /**
			   *document.getElementById("PageOfficeCtrl1").SetWordWaterMarkImage( ImageURL, ImageWidth, IsWashout );
			   *ImageURL  字符串类型，必选参数，图片的路径。
			   *ImageWidth  整数类型，必选参数，图片的宽度（单位：厘米）。如果是0表示采用图片默认宽度。
			   *IsWashout  布尔类型，必选参数，是否冲蚀。true：冲蚀，false：不冲蚀
			   */
			  document.getElementById("PageOfficeCtrl1").SetWordWaterMarkImage("/api/doc/InsertImgWaterMark/logo.jpg", 20, false);
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.AfterDocumentOpened = this.AfterDocumentOpened;
	    }
	}
</script>

