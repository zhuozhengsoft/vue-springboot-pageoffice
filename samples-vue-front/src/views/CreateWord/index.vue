<template>
<div class="Word">
	<div class="zz-content mc clearfix pd-28">
	    <div align="center">
	        <h2>PageOffice 两种创建新文档的方式</h2>
	    </div>
	    <div class="demo mc" style="text-align:left;">
			<hr>
	        <h4 align="center">新建文件</h4></br>
			<h4 align="center">利用属性WebCreateNew创建新文件<input type="button" value="新建文件" @click="creatWord()" /></h4>
			<hr>
	    </div>
	    <div align="center">
	        <h3 align="center" class="fs-12">文档列表</h3>
	        <table class="zz-talbe">
	            <thead>
	            <tr>
	                <th width="22%">
	                    文档名称
	                </th>
	                <th width="10%">
	                    创建日期
	                </th>
	            </tr>
	            </thead>
	            <tbody>
	            <tr v-for="doc in docList" v-model="docList">
	                <td>
						<!-- <input type="button" v-model="doc.subject" value="doc.subject" @click="Word(doc.id)" /> -->
						<a v-model="doc.subject" value="doc.subject" @click="Word(doc.id)">{{doc.subject}}</a>
	                </td>
	                <td>{{doc.submitTime}}</td>
	            </tr>
	            </tbody>
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
        message: ' ',
        docList: {},

      }
    },
    created: function(){
      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
      axios.post("/api/CreateWord/index").then((response) => {
        this.docList = response.data.list;

      }).catch(function (err) {
        console.log(err)
      })
    },
    methods:{
      //控件中的一些常用方法都在这里调用，比如保存，打印等等
      Word(id){
		  POBrowser.openWindowModeless('Word?id='+id, 'width=1200px;height=800px;');
	  },
	  creatWord(){
		  POBrowser.openWindowModeless('createWord', 'width=1200px;height=800px;');
		  
	  },
	  freshIndex(){
		  axios.post("/api/CreateWord/index").then((response) => {
		    this.docList = response.data.list;
		  
		  }).catch(function (err) {
		    console.log(err)
		  })
	  }
    },
    mounted: function(){
      // 将vue中的方法赋值给window
	  window.freshIndex = this.freshIndex;
    }
}
</script>
