<template>
	<div class="Word">
		<div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc;
		        padding: 5px;">
		    <span style="color: red;">操作说明：</span>手动定位光标到当前文档中要插入分页符的位置，然后点“插入分页符”的按钮。<br/>
		    关键代码：看js函数<span style="background-color: Yellow;">InsertPageBreak()</span>
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
	      axios.post("/api/InsertPageBreak/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  InsertPageBreak() {
			  document.getElementById("PageOfficeCtrl1").RunMacro("AddPage", "sub AddPage() \r\n Application.Selection.InsertBreak(7) \r\n End Sub");
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.InsertPageBreak = this.InsertPageBreak;
	    }
	}
</script>

