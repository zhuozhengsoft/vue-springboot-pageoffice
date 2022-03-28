<template>
	<div class="Excel">
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
	      axios.post("/api/InsertSeal/Excel/AddSeal/Excel4").then((response) => {
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

			InsertSeal() {
				try {
					document.getElementById("PageOfficeCtrl1").ZoomSeal.AddSeal("", "");//第二个参数为空字符串时，盖完章的文档是不受保护的，即盖完章后还可以进行编辑等操作；第二个参数为null时，盖完章的文档是受保护的，即盖完章后不能再进行编辑等操作。
				} catch (e) {
				}
			},

			VerifySeal() {
				document.getElementById("PageOfficeCtrl1").ZoomSeal.VerifySeal();
			}
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.Save = this.Save;
			window.InsertSeal = this.InsertSeal;
			window.VerifySeal = this.VerifySeal;
	    }
	}
</script>

