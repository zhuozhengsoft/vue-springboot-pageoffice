<template>
	<div class="Word">
		<input type="button" value="设置父窗口Count的值加1" @click="increaseCount(1);"/>
		<input type="button" value="设置父窗口Count的值加5，并关闭窗口" @click="increaseCountAndClose(5);"/></br>
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
			count: 0,
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/CallParentFunction/Word?id=123").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  Close() {
			  window.external.close();
		  },
	  
		  increaseCount(value) {
			  var sResult = window.external.CallParentFunc("updateCount(" + value + ");");
			  if (sResult == 'poerror:parentlost') {
				  alert('父页面关闭或跳转刷新了，导致父页面函数没有调用成功！');
				  return false;
			  }
			  document.getElementById("PageOfficeCtrl1").Alert("现在父窗口Count的值为：" + sResult);
		  },
	  
		  increaseCountAndClose(value) {
			  var sResult = window.external.CallParentFunc("updateCount(" + value + ");");
			  if (sResult == 'poerror:parentlost') {
				  alert('父页面关闭或跳转刷新了，导致父页面函数没有调用成功！');
			  }
			  window.external.close();
		  },

	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
	      window.Close = this.Close;
	    }
	}
</script>

