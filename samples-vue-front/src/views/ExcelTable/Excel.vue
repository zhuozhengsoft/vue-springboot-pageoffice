<template>
	<div class="Excel">
		此表格中的数据是使用后台程序填充进去的，请查看Table的代码，使用的OpenTable的方法，可以实现行增长，还可以循环使用原模板Table区域（B4:F13）单元格样式。
	  <div style="height: 800px; width: auto" v-html="poHtmlCode" />
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'Excel',
	    data(){
	      return {
	        poHtmlCode: '',
	
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/ExcelTable/Excel").then((response) => {
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


