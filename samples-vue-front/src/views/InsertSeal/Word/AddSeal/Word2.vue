<template>
	<div class="Word">
		不弹出用户名、密码输入框盖章
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
	      axios.post("/api/InsertSeal/Word/AddSeal/Word2").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  //保存
		  Save() {
			  document.getElementById("PageOfficeCtrl1").WebSave();
		  },
	  
		  //加盖印章
		  InsertSeal() {
			  try {
				  document.getElementById("PageOfficeCtrl1").ZoomSeal.AddSeal("李志");
			  } catch (e) {
			  }
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.Save = this.Save;
			window.InsertSeal = this.InsertSeal;
		
	    }
	}
</script>

