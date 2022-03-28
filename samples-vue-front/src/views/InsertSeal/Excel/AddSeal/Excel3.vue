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
	      axios.post("/api/InsertSeal/Excel/AddSeal/Excel3").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
			    Save() {
			        document.getElementById("PageOfficeCtrl1").WebSave();
			    },
			
			    
			    InsertSeal() {
					var strSealName = "测试公章";//加盖的指定印章名称
			        try {
			            //第一个参数，必填项，标识印章名称（当存在重名的印章时，默认选取第一个印章），且该印章签章人的签章类型方式必须为用户名+密码；第二个参数，可选项，标识是否保护文档，为null时保护文档，为空字符串时不保护文档；第三个参数，可选项，标识盖章指定位置名称，须为英文或数字，不区分大小写
			            document.getElementById("PageOfficeCtrl1").ZoomSeal.AddSealByName(strSealName, null);
			        } catch (e) {
			        }
			    },
			
			    DeleteSeal() {
					var strSealName = "测试公章";//加盖的指定印章名称
			        var iCount = document.getElementById("PageOfficeCtrl1").ZoomSeal.Count;//获取加盖的印章数量
			        var strTempSealName = "";
			        if (iCount > 0) {
			            for (var i = 0; i < iCount; i++) {
			                strTempSealName = document.getElementById("PageOfficeCtrl1").ZoomSeal.Item(i).SealName;//获取加盖的印章名称
			                if (strTempSealName == strSealName) {
			                    document.getElementById("PageOfficeCtrl1").ZoomSeal.Item(i).DeleteSeal();//删除印章
			                    alert("成功删除了“" + strSealName + "”");
			                    break;
			                }
			            }
			        } else {
			            alert("请先在文档中加盖印章后，再执行删除操作。");
			        }
			    },
			
			    ChangePsw() {
			        document.getElementById("PageOfficeCtrl1").ZoomSeal.ShowSettingsBox();
			    }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.Save = this.Save;
			window.InsertSeal = this.InsertSeal;
			window.ChangePsw = this.ChangePsw;
			window.DeleteSeal = this.DeleteSeal;
	    }
	}
</script>

