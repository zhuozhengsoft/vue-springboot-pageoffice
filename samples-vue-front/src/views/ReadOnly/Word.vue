<template>
<div class="Word">
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
      axios.post("/api/ReadOnly/Word").then((response) => {
        this.poHtmlCode = response.data;

      }).catch(function (err) {
        console.log(err)
      })
    },
    methods:{
      //控件中的一些常用方法都在这里调用，比如保存，打印等等
      AfterDocumentOpened() {
		  document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(4, false); //禁止另存
		  document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(5, false); //禁止打印
		  document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(6, false); //禁止页面设置
		  document.getElementById("PageOfficeCtrl1").SetEnableFileCommand(8, false); //禁止打印预览
      }
    },
    mounted: function(){
      // 将vue中的方法赋值给window
	  window.AfterDocumentOpened = this.AfterDocumentOpened;
    }
}
</script>
