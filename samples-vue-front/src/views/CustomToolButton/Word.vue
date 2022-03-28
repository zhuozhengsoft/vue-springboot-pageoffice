<template>
<div class="Word">
	    点击自定义工具栏中的“测试按钮”查看效果。<br/>
	    <img src="/static/images/addbutton.jpg"/>
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
      axios.post("/api/CustomToolButton/Word").then((response) => {
        this.poHtmlCode = response.data;

      }).catch(function (err) {
        console.log(err)
      })
    },
    methods:{
      //控件中的一些常用方法都在这里调用，比如保存，打印等等
      myTest() {
        document.getElementById("PageOfficeCtrl1").Alert("测试成功！");
	  }
    },
    mounted: function(){
      // 将vue中的方法赋值给window
      window.myTest = this.myTest;
    }
}
</script>
