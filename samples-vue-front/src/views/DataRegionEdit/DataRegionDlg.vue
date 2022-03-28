<template>
	<div class="Word">
	 <div style="width: 470px; height: 320px;">
	     <div style="float: left;font-size:12px;">
	         <span>待添加标签：</span><br/>
	         <input id="Text1" v-model="Text1" type="text" @propertychange="searchBookMark(Text1);"
	                @input="searchBookMark(Text1);"/>
	         <input id="Button1" type="button" value="搜索" @click="Button1_onclick()"/>
	         <div style="width: 230px; height: 300px; border: solid 1px #ccc; overflow-y: scroll; ">
	             <table id="bkmkTable" style=" font-size:12px;">
	             </table>
	         </div>
	     </div>
	     <div style="float: right;font-size:12px;">
	         <span>已添加标签：</span><br/>
	         <input id="Text2" v-model="Text2" type="text" @propertychange="searchBookMark2(Text2);"
	                @input="searchBookMark2(Text2);"/>
	         <input id="Button2" type="button" value="搜索" @click="Button2_onclick()"/>
	         <div style="width: 230px; height: 300px; border: solid 1px #ccc; overflow-y: scroll; ">
	             <table id="bkmkTable2" style=" font-size:12px;">
	             </table>
	         </div>
	     </div>
	 </div>
	      
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'Word',
	    data(){
	      return {
	        names: '',
			conts: '',
			bkmkArr: '',
			bkContArr: '',
			addBkmkArr: '',
			addBkmkContArr: '',
			Text1:''
	      }
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
  
		  //首次加载数据
			// load() {
			//  pcheckBks();
			//  searchBookMark('');
			//  searchBookMark2('');
			//  return;
			// },
  
		  //遍历Word文件中已存在的书签名称及其对应的文本内容
		  pcheckBks() {
			  var pBkNames = window.external.CallParentFunc("checkBkNames", "");
			  var pBkConts = window.external.CallParentFunc("checkBkConts", "");
  
			  if (pBkNames != "" && pBkNames != null && pBkConts != "" && pBkConts != null) {
				  var currBkmArr = pBkNames.split(",");
				  var currBkContArr = pBkConts.split(",");
				  var start = 0;
				  // addBkmkArr = new Array(0);
				  // addBkmkContArr = new Array(0);
  
				  for (var i = 0; i < currBkmArr.length; i++) {
					  if (currBkmArr[i] != "") {
						  //alert(currBkmArr[i]);
						  this.addBkmkArr.splice(0, 0, currBkmArr[i]); //添加
						  this.addBkmkContArr.splice(0, 0, currBkContArr[i]); //添加
						  for (var j = 0; j < this.bkmkArr.length; j++) {
							  if (this.bkmkArr[j] == currBkmArr[i]) {
								  start = j;
							  }
						  }
						  this.bkmkArr.splice(start, 1); //删除
						  this.bkContArr.splice(start, 1); //删除
					  }
				  }
			  }
		  },
  
		  //加载左侧数据列表
		  searchBookMark(s) {
			  var tb1 = document.getElementById("bkmkTable");
			  var rCount = tb1.rows.length;
			  for (var i = 0; i < rCount; i++) {
				  tb1.deleteRow(0);
			  }
  
			  var oTable = document.getElementById("bkmkTable");
			  for (var i = 0; i < this.bkmkArr.length; i++) {
				  if (this.bkmkArr[i] != null && this.bkmkArr[i] != "" && 0 == this.bkmkArr[i].toLocaleLowerCase().indexOf(s.toLocaleLowerCase())) {
					  var oTr = oTable.insertRow();
					  var oTd = oTr.insertCell();
					  oTd.innerHTML = this.bkmkArr[i];
					  oTd = oTr.insertCell();
					  oTd.innerHTML = this.bkContArr[i];
					  oTd = oTr.insertCell();
					  oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"add('" + this.bkmkArr[i] + "','" + this.bkContArr[i] + "')\">&nbsp;添加</a>";
				  }
			  }
		  },
  
		  //加载右侧数据列表
		  searchBookMark2(s) {
			  //删除所有行
			  var tb1 = document.getElementById("bkmkTable2");
			  var rCount = tb1.rows.length;
			  for (var i = 0; i < rCount; i++) {
				  tb1.deleteRow(0);
			  }
  
			  var oTable = document.getElementById("bkmkTable2");
			  if (this.addBkmkArr != null) {
				  for (var i = 0; i < this.addBkmkArr.length; i++) {
					  if (this.addBkmkArr[i] != null && 0 == this.addBkmkArr[i].toLocaleLowerCase().indexOf(s.toLocaleLowerCase()) && "" != this.addBkmkArr[i]) {
						  var oTr = oTable.insertRow();
						  var oTd = oTr.insertCell();
						  oTd.innerHTML = this.addBkmkArr[i];
						  oTd = oTr.insertCell();
						  oTd.innerHTML = this.addBkmkContArr[i];
						  oTd = oTr.insertCell();
						  oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"locate('PO_" + this.addBkmkArr[i] + "')\">&nbsp;定位</a>";
						  oTd = oTr.insertCell();
						  oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"del('" + this.addBkmkArr[i] + "','" + this.addBkmkContArr[i] + "')\">&nbsp;删除</a>";
					  }
				  }
			  }
		  },
			//搜索待添加书签
		  Button1_onclick() {
			  var s = document.getElementById("Text1").value.toLocaleLowerCase();
			  var tb1 = document.getElementById("bkmkTable");
			  var rCount = tb1.rows.length;
			  for (var i = 0; i < rCount; i++) {
				  tb1.deleteRow(0);
			  }
  
			  var oTable = document.getElementById("bkmkTable");
			  for (var i = 0; i < this.bkmkArr.length; i++) {
				  if (this.bkmkArr[i] != null && this.bkmkArr[i] != "" && this.bkmkArr[i].toLocaleLowerCase().indexOf(s) >= 0) {
  
					  var oTr = oTable.insertRow();
					  var oTd = oTr.insertCell();
					  oTd.innerHTML = this.bkmkArr[i];
					  oTd = oTr.insertCell();
					  oTd.innerHTML = this.bkContArr[i];
					  oTd = oTr.insertCell();
					  oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"add('" + this.bkmkArr[i] + "','" + this.bkContArr[i] + "')\">&nbsp;添加</a>";
				  }
			  }
		  },
			//搜索已添加书签
		  Button2_onclick() {
			  var s = document.getElementById("Text2").value.toLocaleLowerCase();
			  var tb1 = document.getElementById("bkmkTable2");
			  var rCount = tb1.rows.length;
			  for (var i = 0; i < rCount; i++) {
				  tb1.deleteRow(0);
			  }
  
			  var oTable = document.getElementById("bkmkTable2");
			  if (this.addBkmkArr != null) {
				  for (var i = 0; i < this.addBkmkArr.length; i++) {
  
					  if (this.addBkmkArr[i] != null && this.addBkmkArr[i].toLocaleLowerCase().indexOf(s) >= 0 && "" != this.addBkmkArr[i]) {
						  var oTr = oTable.insertRow();
						  var oTd = oTr.insertCell();
						  oTd.innerHTML = this.addBkmkArr[i];
						  oTd = oTr.insertCell();
						  oTd.innerHTML = this.addBkmkContArr[i];
						  oTd = oTr.insertCell();
						  oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"locate('PO_" + this.addBkmkArr[i] + "')\">&nbsp;定位</a>";
						  oTd = oTr.insertCell();
						  oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"del('" + this.addBkmkArr[i] + "','" + this.addBkmkContArr[i] + "')\">&nbsp;删除</a>";
					  }
  
				  }
			  }
		  },
  
		  //******** 书签操作 ************************************************************
			//添加书签
		  add(name, content) {  
			  if ("true" == window.external.CallParentFunc("addBookMark", "PO_" + name + "=" + content)) {
				  var start = 0;
				  for (var i = 0; i < this.bkmkArr.length; i++) {
					  if (this.bkmkArr[i] == name) {
						  start = i;
					  }
				  }
				  this.addBkmkArr.splice(0, 0, name); //向数组第一个位置（0坐标处）添加一个元素
				  this.addBkmkContArr.splice(0, 0, content);
				  this.bkmkArr.splice(start, 1); //删除数组中相应坐标处的一个元素
				  this.bkContArr.splice(start, 1);
  
				  searchBookMark('');
				  searchBookMark2('');
			  }
		  },
			//删除书签
		  del(name, cont) {
			  if ("true" == window.external.CallParentFunc("delBookMark", "PO_" + name)) {
				  var start = 0;
				  for (var i = 0; i < this.addBkmkArr.length; i++) {
					  if (this.addBkmkArr[i] == name) {
						  start = i;
					  }
				  }
  
				  this.bkmkArr.splice(0, 0, name);
				  this.bkContArr.splice(0, 0, cont);
				  this.addBkmkArr.splice(start, 1);
				  this.addBkmkContArr.splice(start, 1);
  
				  searchBookMark('');
				  searchBookMark2('');
			  }
			  else {
				  alert(0)
			  }
		  },
			//定位书签
		  locate(bkName) {
			  window.external.CallParentFunc("locateBK", bkName);
		  }
	    },
		
	    mounted: function(){
			this.Text1 = document.getElementById("Text1").value;
			this.Text2 = document.getElementById("Text2").value;
			this.names = window.external.CallParentFunc("getBkNames", "");
			this.bkmkArr = this.names.split(";");
			this.conts = window.external.CallParentFunc("getBkConts", "");
			this.bkContArr = this.conts.split(";");
			this.addBkmkArr = new Array(0);
			this.addBkmkContArr = new Array(0);
			
			// 将vue中的方法赋值给window
			window.load = this.load;
			window.Save = this.Save;
			window.locate = this.locate;
			window.del = this.del;
			window.add = this.add;
			window.pcheckBks = this.pcheckBks;
			window.searchBookMark = this.searchBookMark;
			window.searchBookMark2 = this.searchBookMark2;
			
			//首次加载数据
			 pcheckBks();
			 searchBookMark('');
			 searchBookMark2('');
   
		  
	    }
	}
</script>

