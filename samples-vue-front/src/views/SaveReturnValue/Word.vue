<template>
	<div class="Word">
		<div style=" font-size:small; color:Red;">
		        <label>键代码：点右键，选择“查看源文件”，看js函数“Save()</label>
		        <br/>document.getElementById("PageOfficeCtrl1").WebSave()//执行保存操作"
		        <br/>document.getElementById("PageOfficeCtrl1").CustomSaveResult//获取返回值保存页面SaveFile代码fs.setCustomSaveResult("ok");定义的返回值
		        <br/>
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
	      axios.post("/api/SaveReturnValue/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      Save(){
	         document.getElementById("PageOfficeCtrl1").WebSave();
			 alert("保存成功，返回值为：" + document.getElementById("PageOfficeCtrl1").CustomSaveResult);
	        },
	      Close() {
	         window.external.close();
	      }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
		  window.Save = this.Save;
		  window.Close = this.Close;
	    }
	}
</script>

