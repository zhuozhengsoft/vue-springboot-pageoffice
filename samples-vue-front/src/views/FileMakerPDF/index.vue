<template>
	<div class="Word">
		<body style="text-align: center;">
		<a  href="javascript:;" @click="convert(1)">1.Word转PDF并下载PDF文件</a><br><br>
		<a  href="javascript:;" @click="convert(2)">2.Word转PDF并打开PDF文件</a>
		<div id="pgImg" style="with:100px;height:100px;margin-top:20px;display: none;"  >
		    正在生成文件，请稍等：<img src="../../../public/images/pg.gif">
		</div>
		</body>
	
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'Word',
	    data(){
	      return {
	        poHtmlCode: '',
			state: ''
	
	      }
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  myFunc(type){
			   alert("文件生成成功！");
			  document.getElementById("pgImg").style.display="none";
			  if("1"==type){
				  //下载生成的pdf文件
				  window.location.href = "/api/FileMakerPDF/download";
			  }else if("2"==type){
				  //打开pdf文件
				  POBrowser.openWindowModeless('OpenPDF' , 'width=1200px;height=800px;');
			  }
  
		  },
		  convert(type) {
			  document.getElementById("pgImg").style.display="block";
			  POBrowser.openWindowModeless('PDF?type='+type , 'width=1px;height=1px;frame=no;');
		  }
		 
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.myFunc = this.myFunc;
	    }
	}
</script>

