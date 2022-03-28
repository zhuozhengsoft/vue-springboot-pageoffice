<template>
	<div class="Word">
		
		<form id="form1">
		    <div id="ProgressBarSide" style="color: Silver; width: 200px; visibility: hidden;
		        position: absolute;  left: 40%; top: 50%; margin-top: -32px">
		        <span style="color: gray; font-size: 12px; text-align: center;">正在生成并打印请稍后...</span><br/>
		        <div id="ProgressBar" style="background-color: Green; height: 16px; width: 0%; border-width: 1px;
		            border-style: Solid;">
		        </div>
		    </div>
			
		    <div style="text-align: center;">
		        <br/>
		        <span style="color: Red; font-size: 12px;">演示：把数据填充到模板中批量生成5个正式的word文件并打印，请点下面的按钮查看效果</span><br/>
		        <input id="Button1" type="button" value="批量生成并打印Word文件" @click="ConvertFiles()"/>
		        <div id="aDiv" style="display: none; color: Red; font-size: 12px;">
		            <span>转换完成，可在后端saveToFile保存的地址中打开文件名为“maker0.doc”到“maker4.doc”的Word文件</span>
		        </div>
		    </div>
		    <div style="width: 1px; height: 1px; overflow: hidden;">
		        <iframe id="iframe1" name="iframe1" src="">
					<div style="height: 800px; width: auto" v-html="poHtmlCode" />
				</iframe>
		    </div>
		</form>
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'Word',
	    data(){
	      return {
	        poHtmlCode: '',
			count: 0,
	      }
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  myFunc() {
			  this.count++;
			  if (this.count < 5) {
				  document.getElementById("iframe1").src = axios.post("/api/PrintFiles/Word?id="+this.count).then((response) => {
															this.poHtmlCode = response.data;
													
														  }).catch(function (err) {
															console.log(err)
														  });
				  // console.log(this.count);
  
				  //设置进度条
				  document.getElementById("ProgressBarSide").style.visibility = "visible";
				  document.getElementById("ProgressBar").style.width = this.count + "0%";
			  } else {
				  //隐藏进度条div
				  document.getElementById("ProgressBarSide").style.visibility = "hidden";
				  this.count = 0;
				  //重置进度条
				  document.getElementById("ProgressBar").style.width = "0%";
				  document.getElementById("aDiv").style.display = "";
				  //alert('批量转换完毕！');
			  }
		  },
  
		  //批量转换完毕
		  ConvertFiles() {
			  //第一次让子页面自刷新
			  let iframe = document.getElementById("iframe1");
			  iframe.src = axios.post("/api/PrintFiles/Word?id="+this.count).then((response) => {
							this.poHtmlCode = response.data;
					
						  }).catch(function (err) {
							console.log(err)
						  });

		  },
		  OnProgressComplete() {
			document.getElementById("FileMakerCtrl1").PrintOut();
			setTimeout(myFunc(), 500 );
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.OnProgressComplete = this.OnProgressComplete;
			window.myFunc = this.myFunc;
	    }
	}
</script>

