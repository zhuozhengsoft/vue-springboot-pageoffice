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
	      axios.post("/api/DisableCopyOut/Word").then((response) => {
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
	  
		  PrintSet() {
			  document.getElementById("PageOfficeCtrl1").ShowDialog(5);
		  },
	  
		  PrintFile() {
			  document.getElementById("PageOfficeCtrl1").ShowDialog(4);
		  },
	  
		  Close() {
			  window.external.close();
		  },
	  
		  IsFullScreen() {
			  document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
		  },
	  
		  //文档关闭前先提示用户是否保存
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
		  window.PrintSet = this.PrintSet;
		  window.PrintFile = this.PrintFile;
		  window.Close = this.Close;
		  window.IsFullScreen = this.IsFullScreen;
		  window.BeforeBrowserClosed = this.BeforeBrowserClosed;
	    }
	}
</script>

