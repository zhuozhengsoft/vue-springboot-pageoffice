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
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/PDFSearch/PDF").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	     SearchText() {
			 document.getElementById("PDFCtrl1").SearchText();
		 },
	 
		 SearchTextNext() {
			 document.getElementById("PDFCtrl1").SearchTextNext();
		 },
	 
		 SearchTextPrev() {
			 document.getElementById("PDFCtrl1").SearchTextPrev();
		 },
	 
		 SetPageReal() {
			 document.getElementById("PDFCtrl1").SetPageFit(1);
		 },
	 
		 SetPageFit() {
			 document.getElementById("PDFCtrl1").SetPageFit(2);
		 },
	 
		 SetPageWidth() {
			 document.getElementById("PDFCtrl1").SetPageFit(3);
		 } 
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.SearchText = this.SearchText;
			window.SearchTextNext = this.SearchTextNext;
			window.SearchTextPrev = this.SearchTextPrev;
			window.SetPageReal = this.SetPageReal;
			window.SetPageFit = this.SetPageFit;
			window.SetPageWidth = this.SetPageWidth;
	    }
	}
</script>

