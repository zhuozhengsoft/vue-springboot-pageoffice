<template>
	<div class="Word">
		“删除行”按钮可以删除光标所在行
	  <div style="height: 800px; width: auto" v-html="poHtmlCode" />
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'Word',
	    data(){
	      return {
	        poHtmlCode: '',
	
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/DeleteRow/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  DeleteRow() {
			  var mac = "Function myfunc()" + " \r\n"
				  + "Application.Selection.HomeKey Unit:=wdLine " + " \r\n"
				  + "Application.Selection.EndKey Unit:=wdLine, Extend:=true " + " \r\n"
				  + "Application.Selection.Cells.Delete ShiftCells:=wdDeleteCellsEntireRow " + " \r\n"
				  + "Application.Selection.TypeBackspace " + " \r\n"
				  + "End Function " + " \r\n";
			  document.getElementById("PageOfficeCtrl1").RunMacro("myfunc", mac);
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.DeleteRow = this.DeleteRow;
	    }
	}
</script>

