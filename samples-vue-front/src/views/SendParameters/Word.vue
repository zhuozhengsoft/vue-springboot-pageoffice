<template>
	<div class="Word">
		<div style="font-size: 14px;">
		        <div style="border:1px solid black;">PageOffice给保存页面传值的三种方式：</br>
		            <span style="color: Red;">1.通过设置保存页面的url中的?给保存页面传递参数：</span></br>
		            &nbsp;&nbsp;&nbsp;例:pocCtrl.setSaveFilePage("SaveFile.jsp?id=1");</br>
		            &nbsp;&nbsp;&nbsp;保存页面获取参数的方法：</br>
		            &nbsp;&nbsp;&nbsp;String value=request.getParameter("id");</br>
		
		            <span style="color: Red;">2.通过input隐藏域给保存页面传递参数：</span></br>
		            &nbsp;&nbsp;&nbsp;例:
		            <xmp><input id="age" name="age" type="hidden" value="25"/></xmp>
		            </br>
		            &nbsp;&nbsp;&nbsp;保存页面获取参数的方法：</br>
		            &nbsp;&nbsp;&nbsp;String age=fs.getFormField("age");</br>
		            &nbsp;&nbsp;&nbsp;注意：获取Form控件传递过来的参数值，fs.getFormField("参数名")方法中的参数名是当前控件的“id”属性，更多详细代码请查看SaveFile.jsp。<br>
		
		            <span style="color: Red;">3.通过Form控件给保存页面传递参数(这里的Form控件包括输入框、下拉框、单选框、复选框、TextArea等类型的控件)：</span></br>
		            &nbsp;&nbsp;&nbsp;例:
		            <xmp><input id="userName" name="userName" type="text" value="张三"/></xmp>
		            </br>
		            &nbsp;&nbsp;&nbsp;保存页面获取参数的方法：</br>
		            &nbsp;&nbsp;&nbsp;String name=fs.getFormField("userName");</br>
		            &nbsp;&nbsp;&nbsp;注意：获取Form控件传递过来的参数值，fs.getFormField("参数名")方法中的参数名是当前控件的“id”属性，更多详细代码请查看SaveFile.jsp。<br>
		        </div>
		        <span style="color: Red;">更新人员信息：</span><br/>
		        <input id="age" name="age" type="hidden" value="25"/>
		        <span style="color: Red;">姓名：</span><input id="userName" name="userName" type="text" value="张三"/><br/>
		        <span style="color: Red;">性别：</span>
		        <select id="selSex" name="selSex">
		        <option value="男">男</option>
		        <option value="女">女</option>
		        </select>
		    </div>  
	    <div style="height: 800px; width: auto" v-html="poHtmlCode"></div>
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
	      axios.post("/api/SendParameters/Word").then((response) => {
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

