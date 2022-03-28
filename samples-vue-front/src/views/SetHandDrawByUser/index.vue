<template>
	<div class="Word">
		<div v-if="div1">
			<div class="flow4">
				<input type="button" @click="exit()" value="返回登录页" />
				<strong>当前用户：</strong>
				<span style="color: Red;">{{userName}}</span>
			</div>
			<div style="height: 800px; width: auto" v-html="poHtmlCode"/>
		</div>
		<div v-if="div2">
			<form id="form1">
				<div style=" text-align:center;">
					<div>请选择登录用户：</div>
					<br/>
					<select name="user" v-model="user">
						<option value="zhangsan">张三</option>
						<option value="lisi">李四</option>
						<option value="wangwu">王五</option>
					</select><br/><br/>
					<input type="button" @click="openWord()" value="打开文件"/><br/><br/>
					<div style=" color:Red;">不同的用户登录后，在文档中隐藏其他人的手写批注，只显示当前用户的手写批注。</div>
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
			userName:''
	      }
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  openWord(){
			  //当点击打开文件时显示div1，隐藏div2
			  this.div1 = true;
			  this.div2 = false;
			  axios.post("/api/SetHandDrawByUser/Word?userName="+this.user).then((response) => {
				this.poHtmlCode = response.data.pageoffice;
						this.userName = response.data.userName;
				
			  }).catch(function (err) {
				console.log(err)
			  })

		  },
		  exit(){
			  this.div1 = false;
			  this.div2 = true;
		  },
		   //保存页面
		  Save() {
			  document.getElementById("PageOfficeCtrl1").WebSave();
		  },

		  //领导圈阅签字
		  StartHandDraw() {
			  document.getElementById("PageOfficeCtrl1").HandDraw.Start();
		  },

		  /*
		  //分层显示手写批注
		  function ShowHandDrawDispBar() {
			  document.getElementById("PageOfficeCtrl1").HandDraw.ShowLayerBar(); ;
		  }*/

		  //全屏/还原
		  IsFullScreen() {
			  document.getElementById("PageOfficeCtrl1").FullScreen = !document.getElementById("PageOfficeCtrl1").FullScreen;
		  },

		  //显示/隐藏用户的手写批注
		  ShowByUserName() {
			  // var userName = document.getElementById("userName").innerText; //从后台获取登录用户名
			  document.getElementById("PageOfficeCtrl1").HandDraw.ShowByUserName(null, false); // 隐藏所有的手写
			  document.getElementById("PageOfficeCtrl1").HandDraw.ShowByUserName(this.userName); // 显示Tom的手写，第二个参数为true时可省略
		  }
		  
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
		  window.Save = this.Save;
		  window.IsFullScreen = this.IsFullScreen;
		  window.StartHandDraw = this.StartHandDraw;
		  window.ShowByUserName = this.ShowByUserName;
	    }
	}
</script>

