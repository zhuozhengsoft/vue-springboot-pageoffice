<template>
	<div class="Word">
		<div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc;
			padding: 5px;">
			注意：<span style="background-color: Yellow;">可通过本地双击打开此docm文件，点击“视图”→“宏”→“查看宏”→“myFunc1”→编辑”，查看myFunc1和myFunc2的具体内容。
			<br/>设置步骤如下：打开一个Word文档，点击“文件”→“选项”→“信任中心”→“信任中心设置”→“宏设置”→勾选上“信任对VBA工程对象模型的访问（V）”</span>
		</div>
		<input id="Button1" type="button" value="执行无返回值宏myFunc1" @click="Button1_onclick()"/>
		<input id="Button2" type="button" value="执行有返回值宏myFunc2" @click="Button2_onclick()"/>
		
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
	      axios.post("/api/RunMacroForDocm/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  Button1_onclick() {
			  /**
			   * document.getElementById("PageOfficeCtrl1").RunMacro( MacroName, MacroScript );.
			   * MacroName	字符串类型，表示宏指令名称。
			   * MacroScript	字符串类型，表示要执行的宏指令代码，可选。
			   */
			  //运行文档中的无返回值的宏命令
			  document.getElementById("PageOfficeCtrl1").RunMacro("myFunc1", "");
		  },
  
		  Button2_onclick() {
			  //运行文档中的有返回值的宏命令
			  var value = document.getElementById("PageOfficeCtrl1").RunMacro("myFunc2", "");
			  document.getElementById("PageOfficeCtrl1").Alert("myFunc2宏的返回值是：" + value);
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window

	    }
	}
</script>

