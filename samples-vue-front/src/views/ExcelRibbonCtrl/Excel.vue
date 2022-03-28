<template>
<div class="Excel">
	<div style="font-size:14px; line-height:20px; ">
		poCtrl.getRibbonBar().setTabVisible("TabHome",false);实现隐藏Ribbon栏中的“开始”命令分组.（"TabHome"为该命令分组的名称,false为隐藏，true为显示）
		<br/>
		poCtrl.getRibbonBar().setSharedVisible("FileSave",
		false);实现隐藏Ribbon快速工具栏中的“保存”按钮.（"FileSave"为按钮的名称,false为隐藏，true为显示）
		<br/>
		poCtrl.getRibbonBar().setGroupVisible("GroupClipboard",
		false);实现隐藏“开始”命令分组中的“剪切板”面板.（"GroupClipboard"为该面板的名称,false为隐藏，true为显示）
		<br/>
		<span style="color:red">注意：此控制是会同步影响到本地文件打开时Ribbon栏中的各个工具按钮的显示状态，当关闭在线编辑窗口时，本地Ribbon栏状态恢复正常。</span>
		<br/><br/>
	</div>
  <div style="width:auto; height:700px;" v-html="poHtmlCode" >
  </div>
</div>
</template>

<script>
  const axios=require('axios');
  export default{
    name: 'Excel',
    data(){
      return {
        message: ' ',
        poHtmlCode: '',

      }
    },
    created: function(){
      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
      axios.post("/api/ExcelRibbonCtrl/Excel").then((response) => {
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
