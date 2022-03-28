<template>
<div class="Word">
	<span style="color: red;">操作说明：请选中文档中一段内容进行测试。</span>
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
      axios.post("/api/OnWordSelectionChange/Word").then((response) => {
        this.poHtmlCode = response.data;

      }).catch(function (err) {
        console.log(err)
      })
    },
    methods:{
      //控件中的一些常用方法都在这里调用，比如保存，打印等等
      OnWordSelectionChange() {
		  var obj = document.getElementById("PageOfficeCtrl1").Document.Application.Selection;
		  if (obj.Range.Text != "") {
			  alert(obj.Range.Text);
		  }
	  }
    },
    mounted: function(){
      // 将vue中的方法赋值给window
	  window.OnWordSelectionChange = this.OnWordSelectionChange;
    }
}
</script>
