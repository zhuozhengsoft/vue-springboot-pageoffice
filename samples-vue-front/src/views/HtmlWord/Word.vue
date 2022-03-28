<template>
	<div class="Word">
		<input type="button" value="打开文档" id="btn1" @click="OpenDocument()"/>
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
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      OpenDocument(){
			  //使打开按钮变灰
			  document.getElementById("btn1").setAttribute("disabled", true);
			  axios.post("/api/HtmlWord/Word?param1=123").then((response) => {
			    this.poHtmlCode = response.data;
			  	
			  }).catch(function (err) {
			    console.log(err)
			  })
		  },
		  Save() {
			  document.getElementById("PageOfficeCtrl1").WebSave();
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window

	    }
	}
</script>

