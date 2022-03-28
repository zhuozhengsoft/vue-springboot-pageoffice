<template>
	<div class="Word">
		
		(1)获取index用？传递过来的id的值：</br>
		输出：id={{id}}</br>
		
		(2)获取父页面文本框中的值：</br>
		输出：{{txt}}</br></br>
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
			txt:'',
			id:''
	      }
	    },
	    created: function(){
			//获取index页面传来的id
			this.id = this.$route.query.id;
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/POBrowserTopic/Word2").then((response) => {
	        this.poHtmlCode = response.data.pageoffice;
			this.txt = response.data.txt;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
				Save() {
					document.getElementById("PageOfficeCtrl1").WebSave();
			
				},
				AfterDocumentOpened() {
					document.getElementById("Text1").value = document.getElementById("PageOfficeCtrl1").DataRegionList.GetDataRegionByName("PO_Title").Value;
				},
			
				setTitleText() {
					document.getElementById("PageOfficeCtrl1").DataRegionList.GetDataRegionByName("PO_Title").Value = document.getElementById("Text1").value;
				}
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.Save = this.Save;
			window.AfterDocumentOpened = this.AfterDocumentOpened;
			window.setTitleText = this.setTitleText;
		
	    }
	}
</script>

