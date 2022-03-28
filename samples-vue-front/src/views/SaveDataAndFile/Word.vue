<template>
<div class="Word">
	    <div>
	        <span style="color: Red; font-size: 14px;">请输入公司名称、年龄、部门等信息后，单击工具栏上的保存按钮</span>
	        <br/>
	        <span style="color: Red; font-size: 14px;">请输入公司名称：</span>
	        <input id="txtName" name="txtCompany" type="text"/>
	        <br/>
	    </div>
  <div style="width:auto; height:700px;" v-html="poHtmlCode" >
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
      axios.post("/api/SaveDataAndFile/Word").then((response) => {
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
       //文档关闭前先提示用户是否保存
		BeforeBrowserClosed() {
				 if (document.getElementById("PageOfficeCtrl1").IsDirty) {
						 if (confirm("提示：文档已被修改，是否继续关闭放弃保存 ？")) {
								 return true;

						 } else {

								 return false;
						 }

				 }

		}
    },
    mounted: function(){
      // 将vue中的方法赋值给window
	  window.Save = this.Save;
      window.BeforeBrowserClosed = this.BeforeBrowserClosed;
    }
}
</script>
