<template>
<div class="Word">
	Word中的"<span style=" color:Red;"> 文件打开了</span>" 是在文档打开的事件中用程序添加进去的。<br/><br/>
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
      axios.post("/api/AfterDocOpened/Word").then((response) => {
        this.poHtmlCode = response.data;

      }).catch(function (err) {
        console.log(err)
      })
    },
    methods:{
      //控件中的一些常用方法都在这里调用，比如保存，打印等等
      AfterDocumentOpened() {
        // 打开文件的时候，给word中当前光标位置赋值一个文本值
        document.getElementById("PageOfficeCtrl1").Document.Application.Selection.Range.Text = "文件打开了";
      }
    },
    mounted: function(){
      // 将vue中的方法赋值给window
      window.AfterDocumentOpened = this.AfterDocumentOpened;
    }
}
</script>
