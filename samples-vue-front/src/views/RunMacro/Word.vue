<template>
	<div class="Word">
		    <div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc;
		        padding: 5px;">
		        注意：<span style="background-color: Yellow;">执行“执行宏myfunc”按钮之前需先设置好MS Word的关于执行宏命令的设置。
		        <br/>设置步骤如下：打开一个Word文档，点击“文件”→“选项”→“信任中心”→“信任中心设置”→“宏设置”→勾选上“信任对VBA工程对象模型的访问（V）”</span>
		    </div>
		    <textarea class="textarea1" id="textarea1" style=" height:87px; width:486px;" rows="" cols="">
					sub myfunc()
				msgbox "123"
				end sub
		    </textarea>
		    <input id="Button1" type="button" value="执行宏myfunc" @click="Button1_onclick()"/>
			
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
	      axios.post("/api/RunMacro/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      Button1_onclick() {
			  document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", document.getElementById("textarea1").value);
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
	      
	    }
	}
</script>



