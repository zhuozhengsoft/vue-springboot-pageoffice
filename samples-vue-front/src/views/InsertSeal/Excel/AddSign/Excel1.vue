<template>
	<div class="Excel">
	  <div style="height: 800px; width: auto" v-html="poHtmlCode" />
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'Excel',
	    data(){
	      return {
	        poHtmlCode: '',
	
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/InsertSeal/Excel/AddSign/Excel1").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
			Save() {
				document.getElementById("PageOfficeCtrl1").WebSave();
				document.getElementById("PageOfficeCtrl1").Alert("文件保存成功。");
			},

			AddHandSign() {
				document.getElementById("PageOfficeCtrl1").ZoomSeal.AddHandSign();
			},

			DeleteHandSign() {
				var iCount = document.getElementById("PageOfficeCtrl1").ZoomSeal.Count;//获取当前文档中加盖的印章数量
				if (iCount > 0) {
					document.getElementById("PageOfficeCtrl1").ZoomSeal.Item(iCount - 1).DeleteSeal();//删除最后一个印章，Item 参数下标从 0 开始
					alert("成功删除了最新的签字。");
				} else {
					alert("请先在文档中签字后，再执行删除操作。");
				}
			},

			VerifySeal() {
				document.getElementById("PageOfficeCtrl1").ZoomSeal.VerifySeal();
			},

			ChangePsw() {
				document.getElementById("PageOfficeCtrl1").ZoomSeal.ShowSettingsBox();
			}
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.Save = this.Save;
			window.AddHandSign = this.AddHandSign;
			window.DeleteHandSign = this.DeleteHandSign;
			window.VerifySeal = this.VerifySeal;
			window.ChangePsw = this.ChangePsw;
	    }
	}
</script>

