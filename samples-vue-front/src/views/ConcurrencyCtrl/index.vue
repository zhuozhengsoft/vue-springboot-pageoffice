<template>
		<div style=" border:solid 1px RoyalBlue; width:600px; text-align:center;  height:300px; margin:100px auto;">
			<div style=" margin-top:50px; height:170px; ">
				<span style=" color:Red;">操作说明：</span><span>首先点击“张三打开文件”会在弹出窗口中打开Word文档（不要关闭窗口），<br/>再点击“李四打开文件”，看并发控制效果。</span><br/><br/>
				1：<a href="javascript:POBrowser.openWindowModeless('Word', 'width=1200px;height=800px;' , '1');">张三打开文件</a><br/><br/>
				2：<a href="javascript:POBrowser.openWindowModeless('Word', 'width=1200px;height=800px;' , '2');">李四打开文件</a>
			</div></br></br></br>
			<div style=" color:Gray;">Copyright &copy 2013 北京卓正志远软件有限公司</div>
		</div>
</template>

<script>

</script>

<style>
</style>
