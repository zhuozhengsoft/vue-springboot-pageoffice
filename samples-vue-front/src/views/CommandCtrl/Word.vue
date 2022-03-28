<template>
	<div class="Word">
		<p>点击“文件”菜单，可以看到“保存”、“另存为”、“页面设置”、“打印”菜单项已经变灰。保存菜单项不仅变灰，Ctrl+S也被禁用。</p>
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
	      axios.post("/api/CommandCtrl/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      AfterDocumentOpened(){
	        document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(3, false); // 禁止保存
	        document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(4, false); // 禁止另存
	        document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(5, false); //禁止打印
	        document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(6, false); // 禁止页面设置
	      }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
	      window.AfterDocumentOpened = this.AfterDocumentOpened;
	    }
	}
</script>


