<template>
	<div class="Word">
		<div><img id="img1"/></div>
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
	      axios.post("/api/SaveFirstPageAsImg/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  Save() {
			  document.getElementById("PageOfficeCtrl1").WebSave();
			  if (document.getElementById("PageOfficeCtrl1").CustomSaveResult == "ok") {
				  document.getElementById("PageOfficeCtrl1").Alert("文档保存成功!");
			  }
		  },
	  
		  SaveFirstAsImg() {
			  document.getElementById("PageOfficeCtrl1").WebSaveAsImage();
			  if (document.getElementById("PageOfficeCtrl1").CustomSaveResult == "ok") {
				  document.getElementById("PageOfficeCtrl1").Alert("图片保存成功!");
				  document.getElementById("img1").src = "/api/doc/SaveFirstPageAsImg/images/test.jpg";
				  document.getElementById("img1").style.width = "200px";
				  document.getElementById("img1").style.height = "290px";
			  }
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.Save = this.Save;
			window.SaveFirstAsImg = this.SaveFirstAsImg;
	    }
	}
</script>

