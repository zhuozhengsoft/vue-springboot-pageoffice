<template>
	<div class="Word">
		<div align="center">
		    <table style="color:red;">
		        <tr>
		            <td>1.可以点"操作"栏中的"编辑"来编辑各个试题</td>
		        </tr>
		        <tr>
		            <td>2.可以选择多个试题，点"生成试卷"按钮，生成试卷</td>
		        </tr>
		    </table>
		</div>
		<div style="text-align: center;">
		    <form id="form1" name="form1" method="post" action="">
		        <table style="background-color: #aaaaff; width: 600px; margin-left:35%; align:center;" v-html="strHtmls"></table>
		        <br/>
		        <input type="button" @click="button1Click();" id="Button1" value="JS方式生成试卷"/><span>（所有版本）</span>
		        <input type="button" @click="button2Click();" id="Button2" value="后台编程生成试卷"/><span
		            style="color:Red;">（专业版、企业版）</span>
		    </form>
		</div>
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'Word',
	    data(){
	      return {
			strHtmls: '',
			pNum:1,
			ids:[],
	      }
	    },
		created: function(){
		  //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
		  axios.post("/api/ExaminationPaper/index").then((response) => {
		    this.strHtmls = response.data;
			
		  }).catch(function (err) {
		    console.log(err)
		  })
		},
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		     // JS方式生成试卷
		     button1Click() {
		         // var ids = "";
		         for (var i = 1; i < 11; i++) {
		             if (document.getElementById("check" + i.toString()).checked) {
		                 this.ids += i.toString() + ",";
		             }
		         }
		         if (this.ids == "")
		             alert("请先选择要生成的试题");
		         else
		             location.href = "javascript:POBrowser.openWindowModeless('Compose?ids=" + this.ids.substr(0, this.ids.length - 1) + "', 'width=1200px;height=800px;');";
		     },
		 
		     // 后台编程生成试卷
		     button2Click() {
		         var ids = "";
		         for (var i = 1; i < 11; i++) {
		             if (document.getElementById("check" + i.toString()).checked) {
		                 ids += i.toString() + ",";
		             }
		         }
		         if (ids == "")
		             alert("请先选择要生成的试题");
		         else
		             location.href = "javascript:POBrowser.openWindowModeless('Compose2?ids=" + ids.substr(0, ids.length - 1) + "', 'width=1200px;height=800px;');";
		     }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.OnProgressComplete = this.OnProgressComplete;
			window.myFunc = this.myFunc;
	    }
	}
</script>

