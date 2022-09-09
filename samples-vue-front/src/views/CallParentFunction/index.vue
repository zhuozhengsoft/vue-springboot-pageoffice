<template>
	<div class="Word">
	  <a href="javascript:POBrowser.openWindowModeless('Word', 'width=1200px;height=800px;scroll=no;');">POBrowser方式打开Word文档</a><br><br><br>
	  <div>Count=<input id="Text1" type="text" value="0"/></div>
	  <router-view></router-view>
	</div>
</template>

<script>
	const axios=require('axios');
	  export default{
	    name: 'Word',
	    data(){
	      return {
	        poHtmlCode: '',
			count: 0,
	      }
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等	      
		  updateCount(value) {
			  this.count = this.count + value;
			  document.getElementById("Text1").value = this.count;
			  return this.count.toString();
		  }
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
	      window.updateCount = this.updateCount;
	    }
	}
</script>

