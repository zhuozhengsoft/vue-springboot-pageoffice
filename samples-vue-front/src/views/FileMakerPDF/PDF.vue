<template>
	<div class="PDF">
		<div style="height: 800px; width: auto" v-html="poHtmlCode" />
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'PDF',
	    data(){
	      return {
	        poHtmlCode: '',
			type: 0,
	
	      }
	    },
	    created: function(){
			this.type = this.$route.query.type;
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/FileMakerPDF/PDF").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  OnProgressComplete() {
			  window.external.CallParentFunc("myFunc("+this.type+");"); //调用父页面的js函数
			  window.external.close();//关闭POBrwoser窗口
		  }
 
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.OnProgressComplete = this.OnProgressComplete;
	    }
	}
</script>

