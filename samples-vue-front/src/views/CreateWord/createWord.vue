<template>
	<div class="Word">
		<div id="textcontent" style="width: 1000px; height: 800px;">
			<div>
				文档主题：
				<input name="FileSubject" id="FileSubject" type="text"
					   onfocus="getFocus()" onblur="lostFocus()" class="boder"
					   style="width: 180px;" value="请输入文档主题"/>
				<input type="button" onclick="Save()" value="保存并关闭"/>
				<input type="button" onclick="Cancel()" value="取消"/>
			</div>
			<div>
				&nbsp;
			</div>
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
        message: ' ',
        poHtmlCode:'',

      }
    },
    created: function(){
      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
      axios.post("/api/CreateWord/creatWord").then((response) => {
        this.poHtmlCode = response.data;

      }).catch(function (err) {
        console.log(err)
      })
    },
    methods:{
      //控件中的一些常用方法都在这里调用，比如保存，打印等等
     Save() {
		 document.getElementById("PageOfficeCtrl1").WebSave();
		 if (document.getElementById("PageOfficeCtrl1").CustomSaveResult == "ok") {
			 alert('保存成功！');
			 //返回列表页面
			 if (window.external.CallParentFunc("freshIndex()") == 'poerror:parentlost') {
				 alert('父页面关闭或跳转刷新了，导致父页面函数没有调用成功！');
			 }
			 window.external.close();//关闭当前POBrower弹出的窗口
		 } else {
			 alert('保存失败！');
		 }
	 },

	 Cancel() {
		 window.external.close();
	 },

	 getFocus() {
		 var str = document.getElementById("FileSubject").value;
		 if (str == "请输入文档主题") {
			 document.getElementById("FileSubject").value = "";
		 }
	 },

	 lostFocus() {
		 var str = document.getElementById("FileSubject").value;
		 if (str.length <= 0) {
			 document.getElementById("FileSubject").value = "请输入文档主题";
		 }
	 },

	 BeforeDocumentSaved() {
		 var str = document.getElementById("FileSubject").value;
		 if (str.length <= 0) {
			 document.getElementById("PageOfficeCtrl1").Alert("请输入文档主题");
			 return false
		 }
		 else
			 return true;
	 },
    },
    mounted: function(){
      // 将vue中的方法赋值给window
	  window.Save = this.Save;
	  window.Cancel = this.Cancel;
	  window.getFocus = this.getFocus;
	  window.lostFocus = this.lostFocus;
	  window.BeforeDocumentSaved = this.BeforeDocumentSaved;
    }
}
</script>
