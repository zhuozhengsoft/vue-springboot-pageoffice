<template>
	<div class="Word">
		<div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc;
		        padding: 5px;">
		        操作：在Html页面添加PageOfficeCtrl控件，再在后台设置PageOfficeCtrl对象的Caption属性<br/>
		        关键代码：<span style="background-color:Yellow;">poCtrl.setCaption("PageOfficeCtrl控件的标题栏内容");//poCtrl为PageOfficeCtrl对象</span>
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
	
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/TitleText/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window

	    }
	}
</script>

