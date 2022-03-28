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
	      axios.post("/api/WordCompare/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
			SaveDocument() {
				document.getElementById("PageOfficeCtrl1").WebSave();
			},
			ShowFile1View() {
				document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.View.ShowRevisionsAndComments = false;
				document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.View.RevisionsView = 1;
			},

			ShowFile2View() {
				document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.View.ShowRevisionsAndComments = false;
				document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.View.RevisionsView = 0;
			},

			ShowCompareView() {
				document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.View.ShowRevisionsAndComments = true;
				document.getElementById("PageOfficeCtrl1").Document.ActiveWindow.View.RevisionsView = 0;
			},

			SetFullScreen() {
				document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
			}
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
	      window.SaveDocument = this.SaveDocument;
		  window.ShowFile1View = this.ShowFile1View;
		  window.ShowFile2View = this.ShowFile2View;
		  window.ShowCompareView = this.ShowCompareView;
		  window.SetFullScreen = this.SetFullScreen;
	    }
	}
</script>

