<template>
	<div class="Word">
		<h4>文件打开页面的token的值为：{{token}}</h4>
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
			token:'',
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
				//在main.js中添加拦截器给请求加上token
	      axios.post("/api/Token/Word").then((response) => {
	        this.poHtmlCode = response.data.pageoffice;
			this.token = response.data.mytoken;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
				Close() {
					window.external.close();
				},
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
					window.Close = this.Close;
	    }
	}
</script>

