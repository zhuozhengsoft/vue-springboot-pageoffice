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
	      axios.post("/api/PromptSave/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  Save() {
			  document.getElementById("PageOfficeCtrl1").WebSave();
		  },
	  
		  CloseFile() {
			  window.external.close();
		  },
	  
		  BeforeBrowserClosed() {
			  if (document.getElementById("PageOfficeCtrl1").IsDirty) {
				  if (confirm("提示：文档已被修改，是否继续关闭放弃保存 ？")) {
					  return true;
	  
				  } else {
	  
					  return false;
				  }
	  
			  }
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.Save = this.Save;
			window.CloseFile = this.CloseFile;
			window.BeforeBrowserClosed = this.BeforeBrowserClosed;
	    }
	}
</script>

