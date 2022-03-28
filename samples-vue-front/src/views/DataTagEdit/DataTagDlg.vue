<template>
	<div class="Word">
		<div style="width: 380px; height: 320px;">
		    <div style="float: left;font-size:12px;">
		        <span>待添加数据标签：</span><br/>
		        <input id="Text1" v-model="Text1" type="text" @propertychange="searchBookMark(Text1);"
		               @input="searchBookMark(Text1);"/>
		        <input id="Button1" type="button" value="搜索" @click="Button1_onclick()"/>
		        <div style="width: 380px; height: 300px; border: solid 1px #ccc; overflow-y: scroll;  ">
		            <table id="tagTable" style=" font-size:12px;">
		            </table>
		        </div>
		    </div>
		
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
			Text1:''
	      }
	    },
	    methods:{
	      
	              //加载数据列表
	              searchBookMark(s) {
					  var names = window.external.CallParentFunc("getTagNames", "");
					  var tagArr = names.split(";");
	                  //删除所有行
	                  var tb1 = document.getElementById("tagTable");
	                  var rCount = tb1.rows.length;
	                  for (var i = 0; i < rCount; i++) {
	                      tb1.deleteRow(0);
	                  }
	      
	                  var oTable = document.getElementById("tagTable");
	                  for (var i = 0; i < tagArr.length; i++) {
	                      if (tagArr[i] != null && tagArr[i] != "" && 0 == tagArr[i].toLocaleLowerCase().indexOf(s.toLocaleLowerCase())) {
	                          var oTr = oTable.insertRow();
	                          var oTd = oTr.insertCell();
	                          oTd.innerHTML = tagArr[i];
	                          oTd = oTr.insertCell();
	                          oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"add('" + tagArr[i] + "','aaaa')\">&nbsp;添加</a>";
	                          oTd = oTr.insertCell();
	                          oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"locate('" + tagArr[i] + "')\">&nbsp;定位</a>";
	                          oTd = oTr.insertCell();
	                          oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"del('" + tagArr[i] + "')\">&nbsp;删除</a>";
	                      }
	                  }
	              },
	      
	      
	              Button1_onclick() {
					  var names = window.external.CallParentFunc("getTagNames", "");
					  var tagArr = names.split(";");
	      
	                  var s = document.getElementById("Text1").value.toLocaleLowerCase();
	                  var tb1 = document.getElementById("tagTable");
	                  var rCount = tb1.rows.length;
	                  for (var i = 0; i < rCount; i++) {
	                      tb1.deleteRow(0);
	                  }
	      
	                  var oTable = document.getElementById("tagTable");
	                  for (var i = 0; i < tagArr.length; i++) {
	                      if (tagArr[i] != null && tagArr[i] != "" && tagArr[i].toLocaleLowerCase().indexOf(s) >= 0) {
	      
	                          var oTr = oTable.insertRow();
	                          var oTd = oTr.insertCell();
	                          oTd.innerHTML = tagArr[i];
	                          oTd = oTr.insertCell();
	                          oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"add('" + tagArr[i] + "')\">&nbsp;添加</a>";
	                          oTd = oTr.insertCell();
	                          oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"locate('" + tagArr[i] + "')\">&nbsp;定位</a>";
	                          oTd = oTr.insertCell();
	                          oTd.innerHTML = "&nbsp;&nbsp;<a href=\"javascript:void(0);\"  onclick= \"del('" + tagArr[i] + "')\">&nbsp;删除</a>";
	                      }
	                  }
	              },
	      
	      
	              //******** Tag 操作 ************************************************************
	      
	              add(name) {
	                  if ("true" == window.external.CallParentFunc("addTag", name)) {
	      
	                  }
	              },
	      
	              del(name) {//alert(name);
	                  if ("false" == window.external.CallParentFunc("delTag", name)) {
	                      alert("请先执行\"定位\"操作，然后再删除。");
	                  }
	              },
	      
	              locate(name) {
	      
	                  window.external.CallParentFunc("locateTag", name);
	              }

		  
	    },
	    mounted: function(){
			this.Text1 = document.getElementById("Text1").value;
			
			// 将vue中的方法赋值给window
			window.del = this.del;
			window.locate = this.locate;
			window.add = this.add;
			window.searchBookMark = this.searchBookMark;
			//首次加载数据
			searchBookMark('');
	      
	      
	    }
	}
</script>

