<template>
	<div class="Word">
		<input name="button" id="Button1" type="button" @click="SetKeyWord(true)" value="高亮显示关键字"/>
		<input name="button" id="Button2" type="button" @click="SetKeyWord(false)" value="取消关键字显示"/>
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
			id:'',
			key:''
	      }
	    },
	    created: function(){
			this.id=this.$route.query.id;
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/SaveAndSearch/Word?id="+this.id).then((response) => {
	        this.poHtmlCode = response.data;
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	         Save() {
	             document.getElementById("PageOfficeCtrl1").WebSave();
	             //document.getElementById("PageOfficeCtrl1").CustomSaveResult获取的是保存页面的返回值
	             if (document.getElementById("PageOfficeCtrl1").CustomSaveResult == "ok")
	                 document.getElementById("PageOfficeCtrl1").Alert("保存成功");
	             else
	                 document.getElementById("PageOfficeCtrl1").Alert(document.getElementById("PageOfficeCtrl1").CustomSaveResult);
	         },
	     
	         SetKeyWord(visible) {
	             if (this.key == "null" || "" == this.key) {
	                 document.getElementById("PageOfficeCtrl1").Alert("关键字为空。");
	                 return;
	             }
	             var sMac = "function myfunc()" + "\r\n"
	                 + "Application.Selection.HomeKey(6) \r\n"
	                 + "Application.Selection.Find.ClearFormatting \r\n"
	                 + "Application.Selection.Find.Replacement.ClearFormatting \r\n"
	                 + "Application.Selection.Find.Text = \"" + this.key + "\" \r\n"
	                 + "While (Application.Selection.Find.Execute()) \r\n"
	                 + "If (" + visible + ") Then \r\n"
	                 + "Application.Selection.Range.HighlightColorIndex = 7 \r\n"
	                 + "Else \r\n"
	                 + "Application.Selection.Range.HighlightColorIndex = 0 \r\n"
	                 + "End If \r\n"
	                 + "Wend \r\n"
	                 + "Application.Selection.HomeKey(6) \r\n"
	                 + "End function";
	     
	             document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", sMac);
	         }
	    },
	    mounted: function(){
			this.key =  decodeURI(window.external.UserParams);
	      // 将vue中的方法赋值给window
			window.Save = this.Save;
	    }
	}
</script>

