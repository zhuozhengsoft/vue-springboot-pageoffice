<template>
	<div class="Word">
		<div>
		    <font color="red">父页面传递过来的参数:</font><input type="text" id="userName" v-model="userName" name="userName"/>
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
			userName: '',
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/GetParentParamValue/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  AfterDocumentOpened() {   
			  this.userName = window.external.UserParams;
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.AfterDocumentOpened = this.AfterDocumentOpened;
	    }
	}
</script>

