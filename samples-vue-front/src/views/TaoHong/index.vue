<template>
	<div class="Word">
	  <div align="center" class="zz-content mc clearfix pd-28">
	      <div class="demo mc">
	          <h2 class="fs-16">
	              pageoffice 在文档流转中的应用
	          </h2>
	          <h3 class="fs-12">
	              演示说明:
	          </h3>
	          <p>
	              提示1：正确运行本演示前，您首先要确定您的机器上已安装Microsoft Office。
	          </p>
	          <p>
	              提示2：如果您是第一次使用PageOffice，在打开文件的时候会提示安装控件，请点允许安装。
	          </p>
	      </div>
	      <div class="zz-talbeBox mc">
	          <h2 class="fs-12">
	              文档列表
	          </h2>
	          <table class="zz-talbe">
	              <thead>
	              <tr @mouseover="onColor(this);" @mouseout="offColor(this);">
	                  <th width="20%" style="text-align:center;">
	                      文档名称
	                  </th>
	                  <th width="20%" style="text-align:center;">
	                      创建日期
	                  </th>
	                  <th width="60%" style="text-align:center;">
	                      操作
	                  </th>
	              </tr>
	              </thead>
	              <tbody>
	              <tr>
	                  <td style="text-align:center;">测试文件</td>
	                  <td style="text-align:center;">2013-05-30</td>
	                  <td style="text-align:center;">
	                      <a href="javascript:POBrowser.openWindowModeless('Word', 'width=1200px;height=800px;');"><span
	                              style=" color:Green;">编辑</span></a>&nbsp;&nbsp;&nbsp;
	                      <a href="javascript:POBrowser.openWindowModeless('taoHong' , 'width=1200px;height=800px;');"><span
	                              style=" color:Green;">套红</span></a>&nbsp;&nbsp;&nbsp;
	                      <a href="javascript:POBrowser.openWindowModeless('readOnly' , 'width=1200px;height=800px;');"><span
	                              style=" color:Green;">正式发文</span></a>
	                  </td>
	              </tr>
	              </tbody>
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
	        poHtmlCode: '',
	
	      }
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
	      onColor(dd) {
	                  dd.style.backgroundColor = "#D7FFEE";
	      },
	      
		  offColor(dd) {
			  dd.style.backgroundColor = "";
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
	      
	    }
	}
</script>


