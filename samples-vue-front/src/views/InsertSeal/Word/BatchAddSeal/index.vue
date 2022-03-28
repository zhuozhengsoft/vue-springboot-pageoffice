<template>
	<div class="PDF">
		<div id="div1" v-if="div1" style="height: 800px; width: auto" v-html="poHtmlCode" />
		<div id="div2" v-if="div2" style="margin:100px" align="center">
		    <form id="form1">
				<h3 align="center">文件列表</h3>
		        <table id="table1" style="border-spacing:20px;border:1px;background-color: darkgray; width: 600px;"
		               cellpadding="3" cellspacing="1">
		            <tr>
		                <td><input name="check" type="checkbox" @click="selectall()"/></td>
		                <td>序号</td>
		                <td>文件名</td>
		                <td>编辑</td>
		            </tr>
		            <tr>
		                <td><input name="check" type="checkbox" value="1"/></td>
		                <td>01</td>
		                <td>test1.doc</td>
		                <!-- <td><a href="Word?id=1">编辑</a></td> -->
						<td><input type="button" @click="editWord(1)" value="编辑"/></td>
		            </tr>
		            <tr>
		                <td><input name="check" type="checkbox" value="2"/></td>
		                <td>02</td>
		                <td>test2.doc</td>
		                <td><input type="button" @click="editWord(2)" value="编辑"/></td>
		            </tr>
		            <tr>
		                <td><input name="check" type="checkbox" value="3"/></td>
		                <td>03</td>
		                <td>test3.doc</td>
		                <td><input type="button" @click="editWord(3)" value="编辑"/></td>
		            </tr>
		            <tr>
		                <td><input name="check" type="checkbox" value="4"/></td>
		                <td>04</td>
		                <td>test4.doc</td>
		                <td><input type="button" @click="editWord(4)" value="编辑"/></td>
		            </tr>
		        </table>
		        </br></br>
		        <input type="button" value="批量盖章" @click="getID()"/>
		    </form>
			<!-- left: 40%; top:30%; -->
			<div id="ProgressBarSide" style="color: Silver; width:500px; visibility: hidden;
			        position: relative; margin-top: 22px ;text-align:center;">
			    <span style="color: gray; font-size: 12px; text-align: left;">正在盖章请稍后...</span><br/>
			    <div id="ProgressBar" style="background-color: Green; height: 16px; width: 0%; border-width: 1px;
			            border-style: Solid;">
			    </div>
			</div>
			<div id="aDiv" style="display: none; color: Red; font-size: 15px;text-align:center;">
			    <span>批量盖章完成，可在下面的地址中查看文件，具体的地址为: {{url}}></span></span>
			</div>
			
			<div style="width: 1px; height: 1px; overflow: hidden;">
			    <iframe id="iframe1" name="iframe1" src="">
					<div style="height: 800px; width: auto" v-html="poHtmlCode" />
				</iframe>
			</div>
		</div>
		
		
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'PDF',
	    data(){
	      return {
			div1:false,
			div2 :true,
			poHtmlCode:'',
	        url: '',
			checkit: true,
			count:0,
			strLength:0,
			str:{},
			isAddSealSuc:false
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/InsertSeal/Word/BatchAddSeal/index").then((response) => {
	        this.url = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
			editWord(id){
				//显示pageoffice控件所在div
				this.div1=true;
				//删除当前弹框页面其他div
				this.div2=false;
				axios.post("/api/InsertSeal/Word/BatchAddSeal/Edit?id="+id).then((response) => {
				this.poHtmlCode = response.data;
			  }).catch(function (err) {
				console.log(err)
			  })
			},
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  
		  selectall() {
			  if (this.checkit) {
				  var obj = document.all.check;
				  for (var i = 0; i < obj.length; i++) {
					  obj[i].checked = true;
					  this.checkit = false;
				  }
			  } else {
				  var obj = document.all.check;
				  for (var i = 0; i < obj.length; i++) {
					  obj[i].checked = false;
					  this.checkit = true;
				  }
  
			  }
  
		  },
		  
  
		  getID() {
			  // var str = new Array();
			  var ids = "";
			  //循环获取选中checkbox的值
			  var check = document.getElementsByName("check");
			  for (var i = 0; i < check.length; i++) {
				  if (check[i].checked) {
					  ids += check[i].value + ",";
				  }
			  }
  
			  if (ids == "" || ids == null || ids == "on,") {
				  alert("请至少选择一个要盖章的文档！");
				  return;
  
			  }
			  //去掉最后一个逗号
			  ids = ids.substring(0, ids.length - 1);
			  var i = ids.indexOf("on");
			  if (i == 0 || i > 0) {
				  //判读是否包含"on"，如果包含，则去掉
				  ids = ids.replace("on,", "");
			  }
			  this.str = ids.split(",");
			  if (this.count > 0) {
				  alert("请刷新当前页面重新选择要盖章的文档!");
				  return;
			  }
			  //第一次自刷新
			  document.getElementById("iframe1").src = axios.post("/api/InsertSeal/Word/BatchAddSeal/AddSeal?id="+this.str[this.count]).then((response) => {
														this.poHtmlCode = response.data;
												
													  }).catch(function (err) {
														console.log(err)
													  })
			  // "Convert?id=" + str[count];
			   document.getElementById("ProgressBarSide").style.visibility = "visible";
			   document.getElementById("ProgressBar").style.width = this.count+1+ "0%";
			  this.strLength = this.str.length;
		  },
  
		  convert() {
			  this.count = this.count + 1;
			  if(this.isAddSealSuc){
				  if (this.count < this.strLength) {
				  				  document.getElementById("iframe1").src = axios.post("/api/InsertSeal/Word/BatchAddSeal/AddSeal?id="+this.str[this.count]).then((response) => {
				  														this.poHtmlCode = response.data;
				  												
				  													  }).catch(function (err) {
				  														console.log(err)
				  													  })
				  				  //设置进度条
				  				  document.getElementById("ProgressBarSide").style.visibility = "visible";
				  				  document.getElementById("ProgressBar").style.width = this.count+1+"0%";
				  } else {
				  				 //隐藏进度条div
				  				 document.getElementById("ProgressBarSide").style.visibility = "hidden";
				  				 this.count = 0;
				  				 //重置进度条
				  				 document.getElementById("ProgressBar").style.width = "0%";
				  				 document.getElementById("aDiv").style.display = "";
				  }
			  }else{
				  document.getElementById("aDiv").style.display = "";
				  
				  //隐藏进度条div
				  document.getElementById("ProgressBarSide").style.visibility = "hidden";
				  //重置进度条
				  document.getElementById("ProgressBar").style.width = "0%";
				  alert("第"+this.count+"份文件盖章失败，无法继续盖章！");
				  return false;
			  }
			  
		  },
		  OnProgressComplete() {
			  //延迟加载循环convert()，保证AfterDocumentOpened()和上一次convert()流程走完
			  setTimeout(convert(), 500 );
			  // alert(111);
		  },
		  Save() {
		  	document.getElementById("PageOfficeCtrl1").WebSave();
		  },
		  AfterDocumentOpened() {
			  try {
				  //先定位到印章位置,再在印章位置上盖章
				  document.getElementById("FileMakerCtrl1").ZoomSeal.LocateSealPosition("Seal1");
				 this.isAddSealSuc =  document.getElementById("FileMakerCtrl1").ZoomSeal.AddSealByName("测试公章", null, "Seal1"); //位置名称不区分大小写
			  } catch (e) {
			  }
			  ;
  
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.OnProgressComplete = this.OnProgressComplete;
			window.Save = this.Save;
			window.convert = this.convert;
			window.AfterDocumentOpened = this.AfterDocumentOpened;
	    }
	}
</script>

