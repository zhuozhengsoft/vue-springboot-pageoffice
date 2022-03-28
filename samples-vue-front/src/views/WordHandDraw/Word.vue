<template>
	<div class="Word">
		<div style="font-size: 12px; line-height: 20px; border-bottom: dotted 1px #ccc; border-top: dotted 1px #ccc; padding: 5px;">
		    <span style="color: red;">操作说明：</span>若想提前设置线宽、颜色、笔型、缩放等，可先点击自定义工具栏上的相应按钮，然后点击“开始手写”按钮。在尚未关闭手写工具栏时，点“撤销最近一次手写”按钮，可撤销最近一次的手写；点击“退出手写”按钮，可退出手写；还可点“设置线宽”、“设置颜色”等按钮对手写批注的颜色、线宽等进行再次设置。
		    <br/>
		    关键代码：点右键，选择“查看源文件”，看js函数
		    <br/>
		    <input id="Button3" type="button" value="设置线宽"
		           @click="SetPenWidth()"/>
		    <input id="Button4" type="button" @click="SetPenColor()"
		           value="设置颜色"/>
		    <input id="Button1" type="button" value="撤销最近一次手写"
		           @click="UndoHandDraw()"/>
		    <input id="Button2" type="button" @click="ExitHandDraw()"
		           value="退出手写"/>
		    <span style="background-color: Yellow;"></span>
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
	
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/WordHandDraw/Word").then((response) => {
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
		  },
		  //开始手写
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
		  //访问手写集合
		  GetHandDrawList() {
  
			  var handDrawList = null;
			  var handDraw = null;
			  handDrawList = document.getElementById("PageOfficeCtrl1").HandDraw;
			  handDrawList.Refresh();
			  document.getElementById("PageOfficeCtrl1").Alert("本文档共有 " + handDrawList.Count + " 个手写批示。");
			  var i = 0; //索引从0开始
			  for (i = 0; i < handDrawList.Count; i++) {
				  handDraw = handDrawList.Item(i);
				  handDraw.Locate();
				  document.getElementById("PageOfficeCtrl1").Alert("第" + handDraw.PageNumber + "页" + ", " + handDraw.UserName + ", " + handDraw.DateTime);
			  }
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
		  window.GetHandDrawList = this.GetHandDrawList;
	    }
	}
</script>

