<template>
	<div class="Word">
		<input id="Button1" type="button" value="隐藏/显示 标题栏" @click="Button1_onclick()"/>
		<input id="Button2" type="button" value="隐藏/显示 菜单栏" @click="Button2_onclick()"/>
		<input id="Button3" type="button" value="隐藏/显示 自定义工具栏" @click="Button3_onclick()"/>
		<input id="Button4" type="button" value="隐藏/显示 Office工具栏" @click="Button4_onclick()"/>
		<br/><br/>
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
	      axios.post("/api/JsControlBars/Word").then((response) => {
	        this.poHtmlCode = response.data;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
			
			// 隐藏/显示 标题栏
			Button1_onclick() {
			    var bVisible = document.getElementById("PageOfficeCtrl1").Titlebar;
			    document.getElementById("PageOfficeCtrl1").Titlebar = !bVisible;
			},
			// 隐藏/显示 菜单栏
			Button2_onclick() {
			    var bVisible = document.getElementById("PageOfficeCtrl1").Menubar;
			    document.getElementById("PageOfficeCtrl1").Menubar = !bVisible;
			},
			// 隐藏/显示 自定义工具栏
			Button3_onclick() {
			    var bVisible = document.getElementById("PageOfficeCtrl1").CustomToolbar;
			    document.getElementById("PageOfficeCtrl1").CustomToolbar = !bVisible;
			},
			// 隐藏/显示 Office工具栏
			Button4_onclick() {
			    var bVisible = document.getElementById("PageOfficeCtrl1").OfficeToolbars;
			    document.getElementById("PageOfficeCtrl1").OfficeToolbars = !bVisible;
			}
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window

	    }
	}
</script>


