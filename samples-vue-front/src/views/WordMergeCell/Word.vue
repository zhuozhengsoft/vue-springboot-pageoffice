<template>
	<div class="Word">
		    <div style="font-size:12px; line-height:20px; border-bottom:dotted 1px #ccc;border-top:dotted 1px #ccc; padding:5px;">
		        关键代码：看js函数<span style="background-color:Yellow;">mergeCellRight()和mergeCellDown()</span>
		    </div>
		    <div style=" font-size:small">
		        <input id="Button1" type="button" value="向右合并一个单元格" @click="mergeCellRight()"/>&nbsp;&nbsp;
		        <input id="Button2" type="button" value="向下合并一个单元格" @click="mergeCellDown()"/><br/>
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
	      axios.post("/api/WordMergeCell/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
				//  // 向右合并一个单元格
				//  mergeCellRight() {
				// 	 var docObj = document.getElementById("PageOfficeCtrl1").Document;
				// 	 docObj.Tables(1).Cell(1, 1).Select();   // 选择单元格（1，1）
				// 	 docObj.Application.Selection.MoveRight(1, 1, 1); // 向右选择一个单元格
				// 	 docObj.Application.Selection.Cells.Merge(); // 合并
				//  }
		  
				//  // 向下合并一个单元格
				//  mergeCellDown() {
				// 	 var docObj = document.getElementById("PageOfficeCtrl1").Document;
				// 	 docObj.Tables(1).Cell(2, 2).Select();   // 选择单元格（2，2）
				// 	 docObj.Application.Selection.MoveDown(5, 1, 1); // 向下移动1个单元格
				// 	 docObj.Application.Selection.Cells.Merge(); // 合并
				//  }
	      
	              // 向右合并一个单元格
	              mergeCellRight() {
	                  var mac = "Function myfunc()" + " \r\n"
	                      + "ActiveDocument.Tables(1).Cell(1, 1).Select " + " \r\n" // 选择单元格（1，1）
	                      + "Application.Selection.MoveRight 1, 1, 1 " + " \r\n"    // 向右选择一个单元格
	                      + "Application.Selection.Cells.Merge " + " \r\n"          // 合并
	                      + "End Function " + " \r\n";
	                  document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", mac);
	              },
	      
	              // 向下合并一个单元格
	              mergeCellDown() {
	                  var mac = "Function myfunc()" + " \r\n"
	                      + "ActiveDocument.Tables(1).Cell(2, 2).Select " + " \r\n"// 选择单元格（2，2）
	                      + "Application.Selection.MoveDown 5, 1, 1 " + " \r\n"    // 向下移动1个单元格
	                      + "Application.Selection.Cells.Merge " + " \r\n"         // 合并
	                      + "End Function " + " \r\n";
	                  document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", mac);
	              }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window

	    }
	}
</script>

