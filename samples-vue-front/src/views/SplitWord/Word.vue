<template>
<div class="Word">
	<div style=" font-size:14px; line-height:20px;">演示说明：<br/>点击“保存”按钮，PageOffice会把文档中三个数据区域（PO_test1，PO_test2，PO_test3）中的内容保存为三个独立的子文件（new1.doc，new2.doc，new3.doc）到“Samples4/SplitWord/doc”
		目录下。
	</div>
	<div style="color: red;font-size:14px; line-height:20px;">
		Word拆分功能只有企业版支持，并且文档的打开模式必须是OpenModeType.docSubmitForm，需要设置数据区域的属性dataRegion1.setSubmitAsFile(true)。<br/><br/>
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
      axios.post("/api/SplitWord/Word").then((response) => {
        this.poHtmlCode = response.data;

      }).catch(function (err) {
        console.log(err)
      })
    },
    methods:{
      //控件中的一些常用方法都在这里调用，比如保存，打印等等
      Save() {
		  document.getElementById("PageOfficeCtrl1").WebSave();
	  }
	  
    },
    mounted: function(){
      // 将vue中的方法赋值给window
	  window.Save = this.Save;
    }
}
</script>
