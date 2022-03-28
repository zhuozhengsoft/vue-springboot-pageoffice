<template>
	<div class="Word">
		<div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc;
		        padding: 5px;">
		    <span style="color: red;">操作说明：</span>选中word文件中的一段文字，然后点“获取选中文字”按钮。<br/>
		    关键代码：看js函数<span style="background-color: Yellow;">getSelectionText()</span>
		</div>
		
		<input id="Button1" type="button" @click="getSelectionText();" value="js获取word选中的文字"/><br/>
		
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
	      axios.post("/api/WordGetSelection/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  getSelectionText() {
			  if (document.getElementById("PageOfficeCtrl1").Document.Application.Selection.Range.Text != "") {
				  document.getElementById("PageOfficeCtrl1").Alert(document.getElementById("PageOfficeCtrl1").Document.Application.Selection.Range.Text);
			  }
			  else {
				  document.getElementById("PageOfficeCtrl1").Alert("您没有选择任何文本。");
			  }
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.getSelectionText = this.getSelectionText;
	    }
	}
</script>

