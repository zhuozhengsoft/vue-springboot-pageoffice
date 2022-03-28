<template>
<div class="Word">
	<div style="color:Red">请导入“/samples-springboot-back/static/doc/ImportWordData”下的ImportWord.doc文档查看演示效果。</div>
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
      axios.post("/api/ImportWordData/Word").then((response) => {
        this.poHtmlCode = response.data;

      }).catch(function (err) {
        console.log(err)
      })
    },
    methods:{
      //控件中的一些常用方法都在这里调用，比如保存，打印等等
      importData() {
		  document.getElementById("PageOfficeCtrl1").WordImportDialog();
	  },
	  submitData() {
		  document.getElementById("PageOfficeCtrl1").WebSave();
	  }
    },
    mounted: function(){
      // 将vue中的方法赋值给window
      window.importData = this.importData;
	  window.submitData = this.submitData;
    }
}
</script>
