<template>
  <div class="Word">
    <div style="height: 800px; width: auto" v-html="poHtmlCode" />
  </div>
</template>
<script>
const axios = require("axios");
export default {
  name: "Word",
  data() {
    return {
      message: "SimpleWord!",
      poHtmlCode: "",
    };
  },
  created: function () {
    //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controllerf方法
    axios
      .post("/api/SimpleWord2/Word")
      .then((response) => {
        this.poHtmlCode = response.data;
      })
      .catch(function (err) {
        console.log(err);
      });
  },
  methods: {
    //控件中的一些常用方法都在这里调用，比如保存，打印等等
    Save() {
      document.getElementById("PageOfficeCtrl1").WebSave();
    }
  },
  mounted: function () {
    // 将PageOffice控件中的方法通过mounted挂载到window对象上，只有挂载后才能被vue组件识别
    window.Save = this.Save;
	
	// 国产操作系统需要加载WPS插件
	if(navigator.userAgent.toLowerCase().indexOf("linux")>0){
		setTimeout(()=>document.getElementById('PageOfficeCtrl1').load('PageOfficeCtrl1','x-wps','59'),1000);
	}
  },
};
</script>

