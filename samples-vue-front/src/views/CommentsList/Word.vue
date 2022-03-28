<template>
<div class="Word">
	    <div style=" width:1300px; height:700px;">
			<div id="Div_Comments" style=" float:left; width:200px; height:700px; border:solid 1px red;">
				<h3>批注列表</h3>
				<input id="Button2" type="button" value="刷新" @click="Button2_onclick()"/>
				<ul id="ul_Comments">

				</ul>
			</div>
			<div style=" width:1050px; height:700px; float:right;">
				<input id="Button1" type="button" value="插入批注" @click="Button1_onclick()"/>
				<input id="Text1" type="text"/>

				<div style="width:auto; height:700px;" v-html="poHtmlCode" ></div>
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
        message: ' ',
        poHtmlCode: '',

      }
    },
    created: function(){
      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
      axios.post("/api/CommentsList/Word").then((response) => {
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
	  refreshList() {
		  var sMac = "Function getComments() " + "\r\n"
			  + "Dim cmts As String " + "\r\n"
			  + "For i=1 To ActiveDocument.Comments.Count " + "\r\n"
			  + "    cmts = cmts +ActiveDocument.Comments.Item(i).Author & \":\" & ActiveDocument.Comments.Item(i).Range.Text + \"||\" " + "\r\n"
			  + "Next" + "\r\n"
			  + "getComments = cmts" + "\r\n"
			  + "End Function ";
  
		  var sComments = document.getElementById("PageOfficeCtrl1").RunMacro("getComments", sMac);
  
		  var arr = sComments.split("||");
  
		  document.getElementById("ul_Comments").innerHTML = "";
		  for (var i = 0; i < arr.length - 1; i++) {
			  document.getElementById("ul_Comments").innerHTML += "<li><a href='#' onclick='goToComment(" + (i + 1) + ")'>" + arr[i] + "</a></li>"
		  }
  
	  },
  
	  getComment(index) {
		  var sMac = "Function getCmtTxt() " + "\r\n"
			  + "getCmtTxt = ActiveDocument.Comments.Item(" + index + ").Range.Text " + "\r\n"
			  + "End Function ";
  
		  return document.getElementById("PageOfficeCtrl1").RunMacro("getCmtTxt", sMac);
	  },
  
	  goToComment(index) {
		  //refreshList();
		  var sMac = "Sub myfunc() " + "\r\n"
			  + "ActiveDocument.Comments.Item(" + index + ").Edit " + "\r\n"
			  + "End Sub ";
  
		  document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", sMac);
  
	  },
  
	  addComment(txt) {
		  var sMac = "Sub myfunc() " + "\r\n"
			  + "Selection.Comments.Add Range:=Selection.Range " + "\r\n"
			  + "Selection.TypeText Text:=\"" + txt + "\" " + "\r\n"
			  + "On Error Resume Next " + "\r\n"
			  + "ActiveWindow.ActivePane.Close " + "\r\n"
			  + "End Sub ";
		  document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", sMac);
	  },
  
	  Button1_onclick() {
		  addComment(document.getElementById("Text1").value);
		  refreshList();
		  document.getElementById("Text1").value = "";
	  },
  
	  Button2_onclick() {
		  refreshList();
	  },
  
	  InsertComment() {
		  document.getElementById("PageOfficeCtrl1").WordInsertComment();
		  var sMac = "Sub myfunc() " + "\r\n"
			  + "On Error Resume Next " + "\r\n"
			  + "ActiveWindow.ActivePane.Close " + "\r\n"
			  + "End Sub ";
		  document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", sMac);
	  }
    },
    mounted: function(){
      // 将vue中的方法赋值给window
	  window.Save = this.Save;
	  window.AfterDocumentOpened = this.AfterDocumentOpened;
	  window.refreshList = this.refreshList;
	  window.getComment = this.getComment;
	  window.goToComment = this.goToComment;
	  window.addComment = this.addComment;
	  window.InsertComment = this.InsertComment;
	  // window.IsFullScreen = this.IsFullScreen;
    }
}
</script>
