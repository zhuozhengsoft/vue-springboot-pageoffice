<template>
	<div class="Word">
		编辑模版 - 在模版中添加盖章位置
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
	      axios.post("/api/InsertSeal/Word/AddSeal/Word4").then((response) => {
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

		InsertSealPos() {
			try {
				document.getElementById("PageOfficeCtrl1").ZoomSeal.AddSealPosition();
			} catch (e) {
			}
			;
		}
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.Save = this.Save;
			window.InsertSealPos = this.InsertSealPos;
	    }
	}
</script>

