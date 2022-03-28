<template>
	<div class="Word">
		    演示了如果使用程序控制数据区域文本的样式。<input type="button" value="点击按钮查看原文件内容样式" @click="OpenWord" /> <br/><br/>
	  <div style="height: 800px; width: auto" v-model="poHtmlCode" v-html="poHtmlCode" />
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
	      axios.post("/api/DataRegionText/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      OpenWord(){
			  axios.post("/api/DataRegionText/Word2").then((response) => {
			    this.poHtmlCode = response.data;
			  	
			  }).catch(function (err) {
			    console.log(err)
			  })
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
	      
	    }
	}
</script>

