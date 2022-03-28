<template>
	<div class="Word">
		<div v-if="div1">
			<div class="flow4">
				<!-- <a href="index"> 返回登录页</a> -->
				<input type="button" @click="exit()" value="返回登录页" />
				<strong>当前用户：</strong>
				<span style="color: Red;">{{strInfo}}</span>
			</div>
			<div style="height: 800px; width: auto" v-html="poHtmlCode"/>
		</div>
		<div v-if="div2">
			<form id="form1">
				<div style=" text-align:center;">
					<div>请选择登录用户：</div>
					<br/>
					<select name="user" v-model="user">
						<option value="zhangsan">A部门经理</option>
						<option value="lisi">B部门经理</option>
					</select><br/><br/>
					<input type="button" @click="openWord()" value="打开文件"/><br/><br/>
					<div style=" color:Red;">不同的用户登录后，在文档中可以编辑的区域不同</div>
				</div>
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
			//控制当前弹框页面显示哪个div，默认显示div2
			div1:false,
			div2:true,
			poHtmlCode:'',
			user:'zhangsan',
			strInfo:''
	      }
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  openWord(){
			  //当点击打开文件时显示div1，隐藏div2
			  this.div1 = true;
			  this.div2 = false;
			  axios.post("/api/SetXlsTableByUser/Excel?userName="+this.user).then((response) => {
				this.poHtmlCode = response.data.pageoffice;
						this.strInfo = response.data.strInfo;
				
			  }).catch(function (err) {
				console.log(err)
			  })

		  },
		  exit(){
			  this.div1 = false;
			  this.div2 = true;
		  },
		  Save() {
		  	document.getElementById("PageOfficeCtrl1").WebSave();
		  },
		  
		  //全屏/还原
		  IsFullScreen() {
		  	document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
		  }
		  
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
		  window.Save = this.Save;
		  window.IsFullScreen = this.IsFullScreen;
	    }
	}
</script>

