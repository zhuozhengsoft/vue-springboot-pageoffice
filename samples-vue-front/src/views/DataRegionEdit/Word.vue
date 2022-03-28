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
	      axios.post("/api/DataRegionEdit/Word").then((response) => {
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
		  ShowDefineDataRegions() {
			  document.getElementById("PageOfficeCtrl1").ShowHtmlModelessDialog("DataRegionDlg", "parameter=xx", "left=300px;top=390px;width=520px;height=410px;");
		  },
		  
		  //获取后台添加的书签名称字符串
		  getBkNames() {
			  var bkNames = document.getElementById("PageOfficeCtrl1").DataRegionList.DefineNames;
			  return bkNames;
		  },
		  //获取后台添加的书签文本字符串
		  getBkConts() {
			  var bkConts = document.getElementById("PageOfficeCtrl1").DataRegionList.DefineCaptions;
			  return bkConts;
		  },
		  //定位书签
		  locateBK(bkName) {
			  var drlist = document.getElementById("PageOfficeCtrl1").DataRegionList;
			  drlist.GetDataRegionByName(bkName).Locate();
			  document.getElementById("PageOfficeCtrl1").Activate();
			  window.focus();
		  },
		  //添加书签
		  addBookMark(param) {
			  var tmpArr = param.split("=");
			  var bkName = tmpArr[0];
			  var content = tmpArr[1];
			  var drlist = document.getElementById("PageOfficeCtrl1").DataRegionList;
			  drlist.Refresh();
			  try {
				  document.getElementById("PageOfficeCtrl1").Document.Application.Selection.Collapse(0);
				  drlist.Add(bkName, content);
				  return "true";
			  } catch (e) {
				  return "false";
			  }
		  },
		  //删除书签
		  delBookMark(bkName) {
			  var drlist = document.getElementById("PageOfficeCtrl1").DataRegionList;
			  try {
				  drlist.Delete(bkName);
				  return "true";
			  } catch (e) {
				  return "false";
			  }
		  },
		  //遍历当前Word中已存在的书签名称
		  checkBkNames() {
			  var drlist = document.getElementById("PageOfficeCtrl1").DataRegionList;
			  drlist.Refresh();
			  var bkName = "";
			  var bkNames = "";
			  for (var i = 0; i < drlist.Count; i++) {
				  bkName = drlist.Item(i).Name;
				  bkNames += bkName.substr(3) + ",";
			  }
			  return bkNames.substr(0, bkNames.length - 1);
		  },
		  //遍历当前Word中已存在的书签文本
		  checkBkConts() {
			  var drlist = document.getElementById("PageOfficeCtrl1").DataRegionList;
			  drlist.Refresh();
			  var bkCont = "";
			  var bkConts = "";
			  for (var i = 0; i < drlist.Count; i++) {
				  bkCont = drlist.Item(i).Value;
				  bkConts += bkCont + ",";
			  }
			  return bkConts.substr(0, bkConts.length - 1);
		  }
	    },
		
	    mounted: function(){
	      // 将vue中的方法赋值给window
	      window.Save = this.Save;
		  window.ShowDefineDataRegions = this.ShowDefineDataRegions;
		  
		  window.getBkNames = this.getBkNames;
		  window.getBkConts = this.getBkConts;
		  window.locateBK = this.locateBK;
		  window.addBookMark = this.addBookMark;
		  window.delBookMark = this.delBookMark;
		  window.checkBkNames = this.checkBkNames;
		  window.checkBkConts = this.checkBkConts;
		  
	    }
	}
</script>

