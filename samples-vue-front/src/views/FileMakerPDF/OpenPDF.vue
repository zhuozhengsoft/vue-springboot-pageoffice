<template>
	<div class="PDF">
		<div style="height: 800px; width: auto" v-html="poHtmlCode" />
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'PDF',
	    data(){
	      return {
	        poHtmlCode: '',
	
	      }
	    },
	    created: function(){
			this.type = this.$route.query.type;
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/FileMakerPDF/OpenPDF").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      AfterDocumentOpened() {
	              //alert(document.getElementById("PDFCtrl1").Caption);
	      },
	      SetBookmarks() {
	      		  document.getElementById("PDFCtrl1").BookmarksVisible = !document.getElementById("PDFCtrl1").BookmarksVisible;
	      },
	      PrintFile() {
	      		  document.getElementById("PDFCtrl1").ShowDialog(4);
	      },
	      SwitchFullScreen() {
	      		  document.getElementById("PDFCtrl1").FullScreen = !document.getElementById("PDFCtrl1").FullScreen;
	      },
	      SetPageReal() {
	      		  document.getElementById("PDFCtrl1").SetPageFit(1);
	      },
	      SetPageFit() {
	      		  document.getElementById("PDFCtrl1").SetPageFit(2);
	      },
	      SetPageWidth() {
	      		  document.getElementById("PDFCtrl1").SetPageFit(3);
	      },
	      ZoomIn() {
	      		  document.getElementById("PDFCtrl1").ZoomIn();
	      },
	      ZoomOut() {
	      		  document.getElementById("PDFCtrl1").ZoomOut();
	      },
	      FirstPage() {
	      		  document.getElementById("PDFCtrl1").GoToFirstPage();
	      },
	      PreviousPage() {
	      		  document.getElementById("PDFCtrl1").GoToPreviousPage();
	      },
	      NextPage() {
	      		  document.getElementById("PDFCtrl1").GoToNextPage();
	      },
	      LastPage() {
	      		  document.getElementById("PDFCtrl1").GoToLastPage();
	      },
	      SetRotateRight() {
	      		  document.getElementById("PDFCtrl1").RotateRight();
	      },
	      SetRotateLeft() {
	      		  document.getElementById("PDFCtrl1").RotateLeft();
	      }
 
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
	      window.AfterDocumentOpened = this.AfterDocumentOpened;
	      window.SetBookmarks = this.SetBookmarks;
	      window.PrintFile = this.PrintFile;
	      window.SwitchFullScreen = this.SwitchFullScreen;
	      window.SetPageReal = this.SetPageReal;
	      window.SetPageFit = this.SetPageFit;
	      window.SetPageWidth = this.SetPageWidth;
	      window.ZoomIn = this.ZoomIn;
	      window.ZoomOut = this.ZoomOut;
	      window.FirstPage = this.FirstPage;
	      window.PreviousPage = this.PreviousPage;
	      window.NextPage = this.NextPage;
	      window.LastPage = this.LastPage;
	      window.SetRotateRight = this.SetRotateRight;
	      window.SetRotateLeft = this.SetRotateLeft;
	    }
	}
</script>

