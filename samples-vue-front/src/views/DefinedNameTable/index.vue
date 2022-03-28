<template>
	<div class="Word">
			<div style="text-align:left;margin-top:30px;">
				<span>1、演示给Excel模板中定义了名称的一块区域赋值，保存时可获取此区域内各单元格的数据：</span><br/><br/>
				<a href="javascript:POBrowser.openWindowModeless('Excel' , 'width=1200px;height=800px;')">openTableByDefinedName方法的简单使用</a>
			</div>
			<div style="text-align:left; margin-top:30px;">
				<span>2、演示如何在不修改代码的情况下，填充数据到用户自定义的两个不同的Excel模板，显示出不同的效果：</span><br/><br/>
				<a href="javascript:POBrowser.openWindowModeless('Excel2?temp=temp1.xls' , 'width=1200px;height=800px;')">自定义Excel模板一显示效果</a><br/><br/>
				<a href="javascript:POBrowser.openWindowModeless('Excel2?temp=temp2.xls' , 'width=1200px;height=800px;')">自定义Excel模板二显示效果</a>
			</div>
			<div style="text-align:left; margin-top:30px;">
				<span>3、演示openTable与openTableByDefinedName方法填充数据到相同Excel模板，在动态填充的数据按实际数据行数自动扩展的情况下，生成的文件显示出不同的效果：</span><br/><br/>
				<a href="javascript:POBrowser.openWindowModeless('Excel6' , 'width=1200px;height=800px;')">打开模板（test4.xls）</a><br/><br/>
				<a href="javascript:POBrowser.openWindowModeless('Excel4' , 'width=1200px;height=800px;')">用openTable填充数据到模板（test4.xls）后的效果――出现表格重叠的问题</a><br/><br/>
				<a href="javascript:POBrowser.openWindowModeless('Excel5' , 'width=1200px;height=800px;')">用openTableByDefinedName填充数据到模板（test4.xls）后的效果――表格互不影响</a>
			</div>
	</div>
</template>

<script>
</script>

<style>
</style>
