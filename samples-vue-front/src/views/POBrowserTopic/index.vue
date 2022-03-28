<template>
	<div class="Word">
	  <div class="zz-content mc clearfix pd-28" align="center">
	      <div class="demo mc">
	          <h3>POBrowser专题</h3>
	      </div>
	  
	      <div style="margin : 10px" align="center">
	          <table style="border-collapse:separate; border-spacing:20px; ">
	              <tr>
	                  <td>1.POBrowser方式最简单的打开编辑保存文档：<a href="javascript:POBrowser.openWindowModeless('Word1' , 'width=1200px;height=800px;');">Word1.html</a>
	                  </td>
	              </tr>
	              <tr>
	                  <td>2.向POBrowser打开的页面传递参数2：<a href="javascript:POBrowser.openWindowModeless('Word2?id=1' , 'width=1200px;height=800px;');" @click="getText()">Word2.html</a>
					  </td>
				  </tr>
	              <tr>
	                  <td>请输入值：
	                      <input type="text" id="ipt" value="1234"></input>
	                  </td>
	              </tr>
	          </table>
	  
	      </div>
	  
	  </div>
	</div>
</template>
<script>
	const axios=require('axios');
	  export default{
	    name: 'Word',
	    data(){
	      return {
	        data: '',
	
	      }
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  getText() {
			  //获取文本框的值，传递给result2
			  var iptValue = document.getElementById("ipt").value;
			  axios.post("/api/POBrowserTopic/Result2?param=" + iptValue).then((response) => {
			    this.data = response.data;
			  	
			  }).catch(function (err) {
			    console.log(err)
			  })
  
		  }
	    }
	}
</script>