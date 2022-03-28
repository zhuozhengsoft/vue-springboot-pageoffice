<template>
	<div class="Word">
		<div class="flow4">
			<a href="#" onclick="window.external.close();">
				<img alt="返回" src="/static/images/return.gif"
					 border="0"/>文件列表</a>&nbsp;&nbsp;<strong>文档主题：</strong>
			<span style="color: Red;">{{subject}}</span> 
			<span style="width: 100px;"></span>
		</div>
		<div style="height: 800px; width: auto" v-html="poHtmlCode" />
	</div>
</template>

<script>
  const axios=require('axios');
  export default{
    name: 'Word',
    data(){
      return {
        message: ' ',
        poHtmlCode:'',
		subject:'',
		id:''
      }
    },
    created: function(){
		this.id = this.$route.query.id;
      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
      axios.post("/api/CreateWord/Word?id="+this.id).then((response) => {
        this.poHtmlCode = response.data.pageoffice;
		this.subject = response.data.subject;
      }).catch(function (err) {
        console.log(err)
      })
    },
    methods:{
      //控件中的一些常用方法都在这里调用，比如保存，打印等等
     Save() {
		 document.getElementById("PageOfficeCtrl1").WebSave();
	 }
    },
    mounted: function(){
      // 将vue中的方法赋值给window
	  window.Save = this.Save;
    }
}
</script>
