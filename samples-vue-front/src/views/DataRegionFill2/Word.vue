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
	      axios.post("/api/DataRegionFill2/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      
	    },
	    mounted: function(){
	      // 国产操作系统需要加载WPS插件
	      if(navigator.userAgent.toLowerCase().indexOf("linux")>0){
	      	setTimeout(()=>document.getElementById('PageOfficeCtrl1').load('PageOfficeCtrl1','x-wps','59'),1000);
	      }
	      
	    }
	}
</script>

