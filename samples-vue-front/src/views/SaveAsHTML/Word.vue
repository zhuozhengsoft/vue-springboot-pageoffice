<template>
<div class="Word">
	<div id="div1"></div>
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
      axios.post("/api/SaveAsHTML/Word").then((response) => {
        this.poHtmlCode = response.data;

      }).catch(function (err) {
        console.log(err)
      })
    },
    methods:{
      //控件中的一些常用方法都在这里调用，比如保存，打印等等
      saveAsHTML() {
        document.getElementById("PageOfficeCtrl1").WebSaveAsHTML();
        document.getElementById("PageOfficeCtrl1").Alert("HTML格式的文件已经保存到 doc/SaveAsHTML 目录下。");
        document.getElementById("div1").innerHTML = "<a href='/api/doc/SaveAsHTML/test.htm?r=" + Math.random() + "'> 查看另存的html文件<a><br><br>";
      }
    },
    mounted: function(){
      // 将vue中的方法赋值给window
      window.saveAsHTML = this.saveAsHTML;
    }
}
</script>
