<template>
	<div class="Word">
		<div class="flow4">
			<a href="#" onclick="window.external.close();"> <img alt="返回" src="../../../public/images/return.gif"
																 border="0"/>文件列表</a>
			<span style="width: 100px;"> </span><strong>文档主题：</strong>
			<span style="color: Red;">测试文件</span>
		</div>
		
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
		  axios.post("/api/TaoHong/readOnly").then((response) => {
		    this.poHtmlCode = response.data;
			
		  }).catch(function (err) {
		    console.log(err)
		  })
		},
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      ShowDialog1() {
			  document.getElementById("PageOfficeCtrl1").ShowDialog(2);
		  },

		  ShowDialog2() {
			  document.getElementById("PageOfficeCtrl1").ShowDialog(5);
		  },

		  ShowDialog3() {
			  document.getElementById("PageOfficeCtrl1").ShowDialog(4);
		  },

		  //全屏/还原
		  IsFullScreen() {
			  document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
	      window.ShowDialog1 = this.ShowDialog1;
		  window.ShowDialog2 = this.ShowDialog2;
		  window.ShowDialog3 = this.ShowDialog3;
		  window.IsFullScreen = this.IsFullScreen;
	    }
	}
</script>


