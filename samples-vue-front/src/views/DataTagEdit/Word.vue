<template>
	<div class="Word">
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
	      axios.post("/api/DataTagEdit/Word").then((response) => {
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
		  ShowDefineDataTags() {
			  document.getElementById("PageOfficeCtrl1").ShowHtmlModelessDialog("DataTagDlg", "parameter=xx", "left=300px;top=390px;width=430px;height=410px;");
  
		  },
		  
		  //获取后台定义的Tag 字符串
		  getTagNames() {
			  var tagNames = document.getElementById("PageOfficeCtrl1").DefineTagNames;
			  return tagNames;
		  },
		  //定位Tag
		  locateTag(tagName) {
			  var appSlt = document.getElementById("PageOfficeCtrl1").Document.Application.Selection;
			  var bFind = false;
			  //appSlt.HomeKey(6);
			  appSlt.Find.ClearFormatting();
			  appSlt.Find.Replacement.ClearFormatting();
  
			  bFind = appSlt.Find.Execute(tagName);
			  if (!bFind) {
				  document.getElementById("PageOfficeCtrl1").Alert("已搜索到文档末尾。");
				  appSlt.HomeKey(6);
			  }
			  window.focus();
		  },
		  //添加Tag
		  addTag(tagName) {
			  try {
				  var tmpRange = document.getElementById("PageOfficeCtrl1").Document.Application.Selection.Range;
				  tmpRange.Text = tagName;
				  tmpRange.Select();
				  return "true";
			  } catch (e) {
				  return "false";
			  }
		  }, 
		  //删除Tag
		  delTag(tagName) {
			  var tmpRange = document.getElementById("PageOfficeCtrl1").Document.Application.Selection.Range;
  
			  if (tagName == tmpRange.Text) {
				  tmpRange.Text = "";
				  return "true";
			  }
			  else
				  return "false";
		  }
		  
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
	      window.Save = this.Save;
		  window.ShowDefineDataTags = this.ShowDefineDataTags;
		  window.getTagNames = this.getTagNames;
		  window.locateTag = this.locateTag;
		  window.addTag = this.addTag;
		  window.delTag = this.delTag;
	    }
	}
</script>

