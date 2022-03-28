<template>
	<div class="Word">
		
		<div style=" width:1380px; height:700px;">
			<div id="Div_Comments" style=" float:left; width:300px; height:700px; border:solid 1px red;">
				<h3>手写批注列表</h3>
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
	      axios.post("/api/HandDrawsList/Word").then((response) => {
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
	  
		  StartHandDraw() {
			  document.getElementById("PageOfficeCtrl1").HandDraw.Start();
		  },
	  
		  //设置线宽
		  SetPenWidth() {
			  document.getElementById("PageOfficeCtrl1").HandDraw.SetPenWidth(5);
		  },
	  
		  //设置颜色
		  SetPenColor() {
	  
			  document.getElementById("PageOfficeCtrl1").HandDraw.SetPenColor(5292104);
		  },
	  
		  //设置笔型
		  SetPenType() {
	  
			  document.getElementById("PageOfficeCtrl1").HandDraw.SetPenType(1);
		  },
	  
		  //设置缩放
		  SetPenZoom() {
	  
			  document.getElementById("PageOfficeCtrl1").HandDraw.SetPenZoom(50);
		  },
	  
		  //撤销最近一次手写
		  UndoHandDraw() {
	  
			  document.getElementById("PageOfficeCtrl1").HandDraw.Undo();
		  },
	  
		  //退出手写
		  ExitHandDraw() {
	  
			  document.getElementById("PageOfficeCtrl1").HandDraw.Exit();
		  },
	  
	  
		  AfterDocumentOpened() {
			  refreshList();
		  },
	  
	  
		  refreshList() {
			  document.getElementById("PageOfficeCtrl1").HandDraw.Refresh();
			  document.getElementById("ul_Comments").innerHTML = "";
			  if (document.getElementById("PageOfficeCtrl1").HandDraw.Count != 0) {
				  for (var i = 0; i < document.getElementById("PageOfficeCtrl1").HandDraw.Count; i++) {
					  var handDraw = document.getElementById("PageOfficeCtrl1").HandDraw.Item(i);
					  var str = " ";
					  str = str + "第" + handDraw.PageNumber + "页" + "," + handDraw.UserName + ", " + handDraw.DateTime;
					  document.getElementById("ul_Comments").innerHTML += "<li><a href='#' onclick='goToHandDraw(" + i + ")'>" + str + "</a></li>";
				  }
			  } else {
				  document.getElementById("PageOfficeCtrl1").Alert("当前文档没有手写批注!");
			  }
	  
		  },
	  
		  //定位到当前手写批注
		  goToHandDraw(index) {
			  document.getElementById("PageOfficeCtrl1").HandDraw.Item(index).Locate();
		  },
	  
		  refresh_click() {
			  //刷新手写批注集合
			  refreshList();
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
		  window.Save = this.Save;
		  window.StartHandDraw = this.StartHandDraw;
		  window.SetPenWidth = this.SetPenWidth;
		  window.SetPenColor = this.SetPenColor;
		  window.SetPenType = this.SetPenType;
		  window.SetPenZoom = this.SetPenZoom;
		  window.UndoHandDraw = this.UndoHandDraw;
		  window.ExitHandDraw = this.ExitHandDraw;
		  window.AfterDocumentOpened = this.AfterDocumentOpened;
		  window.refreshList = this.refreshList;
		  window.goToHandDraw = this.goToHandDraw;
	    }
	}
</script>

