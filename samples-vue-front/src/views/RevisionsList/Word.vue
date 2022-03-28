<template>
	<div class="Word">
		
		<div style=" width:1300px; height:700px;">
		    <div id="Div_Comments" style=" float:left; width:200px; height:700px; border:solid 1px red;">
		        <h3>痕迹列表</h3>
		        <input type="button" name="refresh" value="刷新" @click="refresh_click()"/>
		        <ul id="ul_Comments">
		
		        </ul>
		    </div>
			<div style=" width:1050px; height:700px; float:right;">
		    <div style="height: 800px; width: auto" v-html="poHtmlCode" />
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
	        poHtmlCode: '',
	
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/RevisionsList/Word").then((response) => {
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
		  AfterDocumentOpened() {
			  refreshList();
		  },
		  
		  //获取当前痕迹列表
		  refreshList() {
			  var i;
			  document.getElementById("ul_Comments").innerHTML = "";
			  for (i = 1; i <= document.getElementById("PageOfficeCtrl1").Document.Revisions.Count; i++) {
				  var str = "";
				  str = str + document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(i).Author;
				  var revisionDate = document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(i).Date;
				  //转换为标准时间
				  str = str + " " + dateFormat(revisionDate, "yyyy-MM-dd HH:mm:ss");
	  
				  if (document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(i).Type == "1") {
					  str = str + ' 插入：' + document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(i).Range.Text;
				  }
				  else if (document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(i).Type == "2") {
					  str = str + ' 删除：' + document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(i).Range.Text;
				  }
				  else {
					  str = str + ' 调整格式或样式。';
				  }
				  document.getElementById("ul_Comments").innerHTML += "<li><a href='#' onclick='goToRevision(" + i + ")'>" + str + "</a></li>"
			  }
	  
		  },
	  
		  //GMT时间格式转换为CST
		  dateFormat(date, format) {
			  date = new Date(date);
			  var o = {
				  'M+': date.getMonth() + 1, //month
				  'd+': date.getDate(), //day
				  'H+': date.getHours(), //hour
				  'm+': date.getMinutes(), //minute
				  's+': date.getSeconds(), //second
				  'q+': Math.floor((date.getMonth() + 3) / 3), //quarter
				  'S': date.getMilliseconds() //millisecond
			  };
	  
			  if (/(y+)/.test(format))
				  format = format.replace(RegExp.$1, (date.getFullYear() + '').substr(4 - RegExp.$1.length));
	  
			  for (var k in o)
				  if (new RegExp('(' + k + ')').test(format))
					  format = format.replace(RegExp.$1, RegExp.$1.length == 1 ? o[k] : ('00' + o[k]).substr(('' + o[k]).length));
	  
			  return format;
		  },
	  
		  //定位到当前痕迹
		  goToRevision(index) {
			  document.getElementById("PageOfficeCtrl1").Document.Revisions.Item(index).Range.Select();
		  },
	  
		  //刷新列表
		  refresh_click() {
			  refreshList();
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
		  window.Save = this.Save;
		  window.AfterDocumentOpened = this.AfterDocumentOpened;
		  window.refreshList = this.refreshList;
		  window.dateFormat = this.dateFormat;
		  window.goToRevision = this.goToRevision;
		  // window.Save = this.Save;
	    }
	}
</script>

