<template>
<div class="Word">
	<router-view></router-view>
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
      axios.post("/api/ClickDataRegion/Word").then((response) => {
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
	  IsFullScreen() {
		  document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
	  },
	  OnWordDataRegionClick(Name, Value, Left, Bottom) {
		  if (Name == "PO_deptName") {
			  var strRet = document.getElementById("PageOfficeCtrl1").ShowHtmlModalDialog("HTMLPage", Value, "left=" + Left + "px;top=" + Bottom + "px;width=400px;height=300px;");
			  if (strRet != "") {
				  return (strRet);
			  }
			  else {
				  if ((Value == undefined) || (Value == ""))
					  return " ";
				  else
					  return Value;
			  }
		  }
	  }
    },
    mounted: function(){
      // 将vue中的方法赋值给window
	  window.Save = this.Save;
	  window.IsFullScreen = this.IsFullScreen;
	  window.OnWordDataRegionClick = this.OnWordDataRegionClick;
    }
}
</script>
