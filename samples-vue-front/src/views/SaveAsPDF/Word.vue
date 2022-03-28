<template>
<div class="Word">
	
	<div id="div1">
		<!-- <input type='button' @click='OpenPDF' value='查看另存的pdf文件'/><br><br> -->
	</div>
  <div style="width:auto; height:700px;" v-model="poHtmlCode" v-html="poHtmlCode" >
  </div>
</div>
</template>

<script>
  const axios=require('axios');
  export default{
    name: 'Word',
    data(){
      return {
        message: ' ',
        poHtmlCode: '',

      }
    },
    created: function(){
      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
      axios.post("/api/SaveAsPDF/Word").then((response) => {
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
	  //另存为PDF文件
	  SaveAsPDF() {
		  document.getElementById("PageOfficeCtrl1").WebSaveAsPDF();
		  document.getElementById("PageOfficeCtrl1").Alert("PDF文件已经保存到 SaveAsPDF/doc目录下。");
		  document.getElementById("div1").innerHTML = "<input type='button' id='btn1' onclick='OpenPDF()' value='查看另存的pdf文件'/><br><br>";
	  },
	  //打开pdf
	  OpenPDF(){
		  document.getElementById("btn1").setAttribute("disabled", true);
		  axios.post("/api/SaveAsPDF/PDF?fileName=template.pdf").then((response) => {
		    this.poHtmlCode = response.data;
		  
		  }).catch(function (err) {
		    console.log(err)
		  })
	  },
	  
	  
	  //pdf
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
	  RotateRight() {
	  		  document.getElementById("PDFCtrl1").RotateRight();
	  },
	  RotateLeft() {
	  		  document.getElementById("PDFCtrl1").RotateLeft();
	  }
	  
    },
    mounted: function(){
      // 将vue中的方法赋值给window
	  window.Save = this.Save;
	  window.SaveAsPDF = this.SaveAsPDF;
	  window.OpenPDF = this.OpenPDF;
	  
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
	  window.RotateRight = this.RotateRight;
	  window.RotateLeft = this.RotateLeft;
    }
}
</script>
