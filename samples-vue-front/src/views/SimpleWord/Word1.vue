<template>
	<div class="Word1">
	  <div style="width:auto; height:900px;" v-html="poHtmlCode" >
	  </div>
	</div>
</template>

<script>
  const axios=require('axios');
  export default{
    name: 'Word1',
    data(){
      return {
        message: ' ',
        poHtmlCode: '',

      }
    },
    created: function(){
      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
      axios.post("/api/SimpleWord/Word1").then((response) => {
        this.poHtmlCode = response.data;

      }).catch(function (err) {
        console.log(err)
      })
    },
    methods:{
      //控件中的一些常用方法都在这里调用，比如保存，打印等等
      Save(){
        document.getElementById("PageOfficeCtrl1").WebSave();
      }
    },
    mounted: function(){
      // 将vue中的方法赋值给window
      window.Save = this.Save;
    }
}
</script>
