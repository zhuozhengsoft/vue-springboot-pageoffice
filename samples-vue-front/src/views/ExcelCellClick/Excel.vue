<template>
	<div class="Excel">
		演示：点击Excel单元格弹出自定义对话框的效果。请看实现下面表格中的“部门名称”只能选择的效果。<br/><br/>
	  <div style="height: 800px; width: auto" v-html="poHtmlCode" />
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'Excel',
	    data(){
	      return {
	        poHtmlCode: '',
	
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/ExcelCellClick/Excel").then((response) => {
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
		  OnCellClick(Celladdress, value, left, bottom) {
			  var i = 0;
			  while (i < 5) {//表格第一列的5个单元格都弹出选择对话框
				  if (Celladdress == "$B$" + (4 + i)) {
					  var strRet = document.getElementById("PageOfficeCtrl1").ShowHtmlModalDialog("Select", "", "left=" + left + "px;top=" + bottom + "px;width=320px;height=230px;");
					  if (strRet != "") {
						  return (strRet);
					  }
					  else {
						  if ((value == undefined) || (value == ""))
							  return " ";
						  else
							  return value;
					  }
				  }
				  i++;
			  }
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
	      window.Save = this.Save;
		  window.OnCellClick = this.OnCellClick;
	    }
	}
</script>

