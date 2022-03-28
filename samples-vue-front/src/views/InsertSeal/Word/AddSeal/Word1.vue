<template>
	<div class="Word">
		<span style="color: red;">操作说明：</span>点“加盖印章”按钮即可，插入印章时的用户名为：李志，密码默认为：111111。
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
	      axios.post("/api/InsertSeal/Word/AddSeal/Word1").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		      //保存
		      Save() {
		          document.getElementById("PageOfficeCtrl1").WebSave();
		          document.getElementById("PageOfficeCtrl1").Alert("文件保存成功。");
		      },
		  
		      //加盖印章
		      InsertSeal() {
		          try {
		              document.getElementById("PageOfficeCtrl1").ZoomSeal.AddSeal();
		          } catch (e) {
		          }
		      },
		  
		      //删除印章
		      DeleteSeal() {
		          var iCount = document.getElementById("PageOfficeCtrl1").ZoomSeal.Count;//获取当前文档中加盖的印章数量
		          if (iCount > 0) {
		              document.getElementById("PageOfficeCtrl1").ZoomSeal.Item(iCount - 1).DeleteSeal();//删除最后一个印章，Item 参数下标从 0 开始
		              alert("成功删除了最新加盖的印章。");
		          } else {
		              alert("请先在文档中加盖印章后，再执行删除操作。");
		          }
		      },
		  
		      //验证印章
		      VerifySeal() {
		          document.getElementById("PageOfficeCtrl1").ZoomSeal.VerifySeal();
		      },
		  
		      //修改密码
		      ChangePsw() {
		          document.getElementById("PageOfficeCtrl1").ZoomSeal.ShowSettingsBox();
		      }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.Save = this.Save;
			window.InsertSeal = this.InsertSeal;
			window.DeleteSeal = this.DeleteSeal;
			window.VerifySeal = this.VerifySeal;
			window.ChangePsw = this.ChangePsw;
	    }
	}
</script>

