<template>
	<div class="Word">
		<div style="text-align: center;">
		    <p style="color: Red; font-size: 14px;">
		        1.可以点击“操作”栏中的“编辑”来编辑各个员工的工作条&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br/>
		        2.可以点击“操作”栏中的“查看”来查看各个员工的工作条&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br/>
		        3.可选择多条工资记录，点“生成工资条”按钮，生成工资条&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
		    <table
		            style="width: 60%; font-size: 12px; text-align: center; border: solid 1px #a2c5d9;" v-html="poHtmlCode">
		
		    </table>
		    <br/>
		    <input type="button" value="生成工资条" @click="getID()"/>
		</div>
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
	      axios.post("/api/WordSalaryBill/index").then((response) => {
	        this.poHtmlCode = response.data;
			// alert(this.poHtmlCode)
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      getID() {
			  var ids = "";
			  for (var i = 1; i < 11; i++) {
				  if (document.getElementById("check" + i.toString()).checked) {
					  ids += i.toString() + ",";
				  }
			  }
  
			  if (ids == "")
				  alert("请先选择要生成工资条的记录");
			  else
				  location.href = "javascript:POBrowser.openWindowModeless('Compose?ids=" + ids.substr(0, ids.length - 1) + " ', 'width=1200px;height=800px;');";
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
	      
	    }
	}
</script>

