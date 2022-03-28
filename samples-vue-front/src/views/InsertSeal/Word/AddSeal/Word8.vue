<template>
	<div class="Word">
		    <span style="color: red;">操作说明：</span>点“签字”按钮即可，插入签字时的用户名为：李志，密码默认为：111111。
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
	      axios.post("/api/InsertSeal/Word/AddSeal/Word8").then((response) => {
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
			
				InsertSeal() {
					try {
						document.getElementById("PageOfficeCtrl1").ZoomSeal.AddSeal("", "");
					} catch (e) {
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
			window.InsertSeal = this.InsertSeal;
			window.VerifySeal = this.VerifySeal;
			window.ChangePsw = this.ChangePsw;
	    }
	}
</script>

