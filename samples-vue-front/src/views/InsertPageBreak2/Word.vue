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
	      axios.post("/api/InsertPageBreak2/Word").then((response) => {
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
				 alert("保存成功！请在samples-springboot-back/static/doc/InsertPageBreak2目录下查看合并后的新文档\"test3.doc\"。");
			 }
		 } 
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.Save = this.Save;
	    }
	}
</script>

