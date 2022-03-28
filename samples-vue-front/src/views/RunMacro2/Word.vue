<template>
	<div class="Word">
		<div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc;
			padding: 5px;">
			注意：<span style="background-color: Yellow;">执行“执行宏myFunc”按钮之前需先设置好MS Word的关于执行宏命令的设置。
			<br/>设置步骤如下：打开一个Word文档，点击“文件”→“选项”→“信任中心”→“信任中心设置”→“宏设置”→勾选上“信任对VBA工程对象模型的访问（V）”</span>
		</div>
		
		<textarea id="textarea1" name="textarea1" style=" height:87px; width:486px;" rows="" cols="">
		   Function myFunc1()
		   myFunc1 = "123"
		   End Function
		</textarea>
		
		<input id="Button1" type="button" value="执行宏myFunc1" @click="Button1_onclick()"/>
		<input id="Button2" type="button" value="执行宏myFunc2" @click="RunMacro2()"/>
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
	      axios.post("/api/RunMacro2/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      Button1_onclick() {
			  var value = document.getElementById("PageOfficeCtrl1").RunMacro("myFunc1", document.getElementById("textarea1").value);
			  document.getElementById("PageOfficeCtrl1").Alert("myFunc1宏的返回值是：" + value);
		  },
  
		  RunMacro2() {
			  var value = document.getElementById("PageOfficeCtrl1").RunMacro("myFunc2", 'Function myFunc2() \r\n myFunc2 = "123" \r\n End Function');
			  document.getElementById("PageOfficeCtrl1").Alert(value);
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window

	    }
	}
</script>

