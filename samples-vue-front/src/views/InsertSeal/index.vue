<template>
	<div class="Word">
		<div class="zz-content mc clearfix pd-28" align="center">
			<div class="demo mc">
				<h2 class="fs-16">
					盖章和(手写)签字专题
				</h2>
			</div>
			<div style="margin: 10px" align="center">
				<p>
					点击
					<a href="http://localhost:8086/samples-springboot-back/adminseal.zz" target="_blank">电子印章简易管理平台</a> 添加、删除印章。管理员:admin 密码:111111
				<span  style="font-size:12px;color: red">(为了系统使用印章的安全性，强烈建议修改此登录密码！！)</span>
				</p>
				<p>
					点击右侧的“全部复位”连接，即可重新演示盖章和签字效果： <input type="button" value="全部复位" @click="refreshDemo()" />
				</p>
				<p style="color: red;">
					印章测试用户名：李志，密码：111111或123456
				</p>
			</div>
			<div style="margin: 10px" align="center">
				<h3>
					一、盖章
				</h3>
				<table style="border-collapse: separate; border-spacing: 20px;border: 1px solid #9CF " align="center"
					   width="1200px">
					<tr>
						<th style="border-bottom: 1px solid #9CF ">功能演示</th>
						<th style="border-bottom: 1px solid #9CF ">文件目录</th>
					</tr>
					<tr>
						<td>
							1.常规盖章：
							<a href="javascript:POBrowser.openWindowModeless('Word/AddSeal/Word1','width=1200px;height=800px;');">Word1</a>
						</td>
						<td>
							(Word/AddSeal)
						</td>
					</tr>
					<tr>
						<td>
							2.无需输入用户名、密码盖章：
							<a href="javascript:POBrowser.openWindowModeless('Word/AddSeal/Word2','width=1200px;height=800px;');">Word2</a>
						</td>
						<td>
							(Word/AddSeal)
						</td>
					</tr>
					<tr>
						<td>
							3.无需输入用户名、密码，并且不弹出印章选择框盖章：
							<a href="javascript:POBrowser.openWindowModeless('Word/AddSeal/Word3','width=1200px;height=800px;');">Word3</a>
						</td>
						<td>
							(Word/AddSeal)
						</td>
					</tr>
					<tr>
						<td>
							4.编辑模版 - 在模版中添加盖章位置：
							<a href="javascript:POBrowser.openWindowModeless('Word/AddSeal/Word4','width=1200px;height=800px;');">Word4</a>
						</td>
						<td>
							(Word/AddSeal)
						</td>
					</tr>
					<tr>
						<td>
							5.常规指定位置盖章，加盖印章到模板中的指定位置：
							<a href="javascript:POBrowser.openWindowModeless('Word/AddSeal/Word5','width=1200px;height=800px;');">Word5</a>
						</td>
						<td>
							(Word/AddSeal)
						</td>
					</tr>
					<tr>
						<td>
							6.无需输入用户名、密码加盖印章到模板中的指定位置：
							<a href="javascript:POBrowser.openWindowModeless('Word/AddSeal/Word6','width=1200px;height=800px;');">Word6</a>
						</td>
						<td>
							(Word/AddSeal)
						</td>
					</tr>
					<tr>
						<td>7.无需输入用户名、密码，并且不弹出印章选择框加盖印章到模板中的指定位置：
							<a href="javascript:POBrowser.openWindowModeless('Word/AddSeal/Word7','width=1200px;height=800px;');">Word7</a>
						</td>
						<td>
							(Word/AddSeal)
						</td>
					</tr>
		
					<tr>
						<td>
							8.特殊盖章需求实现：盖章后印章不保护文档内容，用户仍可编辑修改，印章不会出现失效字样：
							<a href="javascript:POBrowser.openWindowModeless('Word/AddSeal/Word8','width=1200px;height=800px;');">Word8</a>
						</td>
						<td>
							(Word/AddSeal)
						</td>
					</tr>
		
					<tr>
						<td>
							9.批量盖章：
							<a href="javascript:POBrowser.openWindowModeless('Word/BatchAddSeal/index','width=1200px;height=800px;');">index</a>
						</td>
						<td>
							(Word/BatchAddSeal)
						</td>
					</tr>
					<tr>
						<td>
							10.删除印章：
							<a href="javascript:POBrowser.openWindowModeless('Word/AddSeal/Word10','width=1200px;height=800px;');">Word10</a>
						</td>
						<td>
							(Word/AddSeal)
						</td>
					</tr>
		
				</table>
				<div style="margin: 10px" align="center">
					<h3>二、(手写)签字</h3>
					<table style="border-collapse: separate; border-spacing: 20px;border: 1px solid #9CF" align="center"
						   width="1200px">
						<tr>
							<th style="border-bottom: 1px solid #9CF ">功能演示</th>
							<th style="border-bottom: 1px solid #9CF ">文件目录</th>
						</tr>
						<tr>
							<td>
								1.常规(手写)签字：
								<a href="javascript:POBrowser.openWindowModeless('Word/AddSign/Word1','width=1200px;height=800px;');">Word1</a>
							</td>
							<td>
								(Word/AddSign)
							</td>
						</tr>
						<tr>
							<td>
								2.无需输入用户名密码签字：
								<a href="javascript:POBrowser.openWindowModeless('Word/AddSign/Word2','width=1200px;height=800px;');">Word2</a>
							</td>
							<td>
								(Word/AddSign)
							</td>
						</tr>
						<tr>
							<td>
								3.(手写)签字到文档指定位置：
								<a href="javascript:POBrowser.openWindowModeless('Word/AddSign/Word3','width=1200px;height=800px;');">Word3</a>
							</td>
							<td>
								(Word/AddSign)
							</td>
						</tr>
						<tr>
							<td>
								4.无需输入用户名、密码，(手写)签字到模板中的指定位置（参考一、4示例在模板中添加签字位置）：
								<a href="javascript:POBrowser.openWindowModeless('Word/AddSign/Word4','width=1200px;height=800px;');">Word4</a>
							</td>
							<td>
								(Word/AddSign)
							</td>
						</tr>
						<tr>
							<td>
								5.特殊(手写)签字需求实现：签批后(手写)签字不保护文档内容，用户仍可编辑修改，(手写)签字不会出现失效字样：
								<a href="javascript:POBrowser.openWindowModeless('Word/AddSign/Word5','width=1200px;height=800px;');">Word5</a>
							</td>
							<td>
								(Word/AddSign)
							</td>
						</tr>
					</table>
				</div>
				<div style="margin: 10px" align="center">
					<h3>三、PDF盖章、签字<span style=" color:Red;">（专业版、企业版）</span></h3>
					<table style="border-collapse: separate; border-spacing: 20px;border: 1px solid #9CF" align="center"
						   width="1200px">
						<tr>
							<th style="border-bottom: 1px solid #9CF; width:85%;">功能演示</th>
							<th style="border-bottom: 1px solid #9CF">文件目录</th>
						</tr>
						<tr>
							<td>
								1.常规盖章：
								<a href="javascript:POBrowser.openWindowModeless('PDF/AddSeal/PDF1','width=1200px;height=850px;');">PDF1</a>
							</td>
							<td>
								(PDF/AddSeal)
							</td>
						</tr>
						<tr>
							<td>
								2.常规签字：
								<a href="javascript:POBrowser.openWindowModeless('PDF/AddSign/PDF1' ,'width=1200px;height=850px;');">PDF2</a>
							</td>
							<td>
								(PDF/AddSign)
							</td>
						</tr>
					</table>
				</div>
		
		
				<div style="margin: 10px" align="center">
					<h3>
						四、Excel盖章
					</h3>
					<table style="border-collapse: separate; border-spacing: 20px;border: 1px solid #9CF " align="center"
						   width="1200px">
						<tr>
							<th style="border-bottom: 1px solid #9CF; width:85%;">功能演示</th>
							<th style="border-bottom: 1px solid #9CF ">文件目录</th>
						</tr>
						<tr>
							<td>
								1.常规盖章：
								<a href="javascript:POBrowser.openWindowModeless('Excel/AddSeal/Excel1','width=1200px;height=850px;');">Excel1</a>
							</td>
							<td>
								(Excel/AddSeal)
							</td>
						</tr>
		
						<tr>
							<td>
								2.无需输入用户名密码盖章：
								<a href="javascript:POBrowser.openWindowModeless('Excel/AddSeal/Excel2','width=1200px;height=800px;');">Excel2</a>
							</td>
							<td>
								(Excel/AddSeal)
							</td>
						</tr>
						<tr>
							<td>
								3.无需输入用户名密码，并且不显示印章选择对话框：
								<a href="javascript:POBrowser.openWindowModeless('Excel/AddSeal/Excel3','width=1200px;height=800px;');">Excel3</a>
							</td>
							<td>
								(Excel/AddSeal)
							</td>
						</tr>
						<tr>
							<td>
								4.特殊盖章需求实现：盖章后印章不保护文档内容，用户仍可编辑修改，印章不会出现失效字样：
								<a href="javascript:POBrowser.openWindowModeless('Excel/AddSeal/Excel4','width=1200px;height=800px;');">Excel4</a>
							</td>
							<td>
								(Excel/AddSeal)
							</td>
						</tr>
						<tr>
							<td>
								5.删除印章：
								<a href="javascript:POBrowser.openWindowModeless('Excel/AddSeal/Excel5','width=1200px;height=800px;');">Excel5</a>
							</td>
							<td>
								(Excel/AddSeal)
							</td>
						</tr>
		
					</table>
				</div>
				<div style="margin: 10px" align="center">
					<h3>四、Excel(手写)签字</h3>
					<table style="border-collapse: separate; border-spacing: 20px;border: 1px solid #9CF" align="center"
						   width="1200px">
						<tr>
							<th style="border-bottom: 1px solid #9CF; width:85%;">功能演示</th>
							<th style="border-bottom: 1px solid #9CF ">文件目录</th>
						</tr>
						<tr>
							<td>
								1.常规(手写)签字：
								<a href="javascript:POBrowser.openWindowModeless('Excel/AddSign/Excel1','width=1200px;height=850px;');">Excel1</a>
							</td>
							<td>
								(Excel/AddSign)
							</td>
						</tr>
						<tr>
							<td>
								2.无需输入用户名密码(手写)签字：
								<a href="javascript:POBrowser.openWindowModeless('Excel/AddSign/Excel2','width=1200px;height=850px;');">Excel2</a>
							</td>
							<td>
								(Excel/AddSign)
							</td>
						</tr>
						<tr>
							<td>
								3.特殊(手写)签字需求实现：(手写)签批后签字不保护文档内容，用户仍可编辑修改，签字不会出现失效字样：
								<a href="javascript:POBrowser.openWindowModeless('Excel/AddSign/Excel3','width=1200px;height=800px;');">Excel3</a>
							</td>
							<td>
								(Excel/AddSign)
							</td>
						</tr>
					</table>
		
				</div>
		
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
	
	      }
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  refreshDemo() {
			  if (confirm("确认复位所有示例的文档吗？")) {
				  
				  axios.post("/api/InsertSeal/refresh").then((response) => {
					alert(response.data);
				  }).catch(function (err) {
				    console.log(err)
				  })
			  }
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window

	    }
	}
</script>