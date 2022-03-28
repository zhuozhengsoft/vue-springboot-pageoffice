<template>
<div class="Word">
	演示: 文档保存前和保存后执行的事件。<br/><br/>
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
      axios.post("/api/BeforeAndAfterSave/Word").then((response) => {
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
        BeforeDocumentSaved() {
            document.getElementById("PageOfficeCtrl1").Alert("BeforeDocumentSaved事件：文件就要开始保存了.");
            return true;
          },
      
        AfterDocumentSaved(IsSaved) {
            if (IsSaved) {
                document.getElementById("PageOfficeCtrl1").Alert("AfterDocumentSaved事件：文件保存成功了.");
            }
        }
    },
    mounted: function(){
      // 将vue中的方法赋值给window
	  window.Save = this.Save;
      window.BeforeDocumentSaved = this.BeforeDocumentSaved;
	  window.AfterDocumentSaved = this.AfterDocumentSaved;
    }
}
</script>
