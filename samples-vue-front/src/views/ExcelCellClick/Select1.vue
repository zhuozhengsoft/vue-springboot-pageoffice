<template>
	<div class="Excel">
		<form id="form1">
			<select v-model="selected">
			  <option v-for="option in options" v-bind:value="option.value">
			   {{ option.text }}
			  </option>
			</select>
			    <button type="button" @click="CheckValue();">确定</button>
			    <button type="button" @click="window.opener=null;window.open('','_self');window.close();">取消</button>
		</form>
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'Excel',
	    data(){
	      return {
	        poHtmlCode: '',
			selected:'行政部',
			options: [
			     { text: '行政部', value: '行政部' },
			     { text: '财务部', value: '财务部' },
			     { text: '市场一部', value: '市场一部' },
				 { text: '市场二部', value: '市场二部' },
				 { text: '市场三部', value: '市场三部' },
				 { text: '市场四部', value: '市场四部' }
			   ]
	      }
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
		  CheckValue(theForm) {
			  window.external.DialogResult = this.selected; //设置ShowHtmlModalDialog的返回值
			  window.opener = null;
			  window.open('', '_self');
			  window.close();
			  return;
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
	     
	    }
	}
</script>

