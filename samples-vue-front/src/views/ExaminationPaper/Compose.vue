<template>
	<div class="Word">
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
					operateStr:'',//后端返回的creat方法
					ids:'',
					idArr:[],
					pNum:1
	      }
	    },
	    created: function(){
			this.ids = this.$route.query.ids;
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/ExaminationPaper/Compose?ids="+this.ids).then((response) => {
	        this.poHtmlCode = response.data.pageoffice;
			this.idArr = response.data.idArr;
	
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
			Create(){
				var obj = document.getElementById('PageOfficeCtrl1').Document.Application;
				obj.Selection.EndKey(6);
				for(var i=0; i<this.idArr.length; i++) {
					 obj.Selection.TypeParagraph();
					 obj.Selection.Range.Text = "'"+ this.pNum + ".'";
					 obj.Selection.EndKey(5,1);
					 obj.Selection.MoveRight(1,1);
					 document.getElementById('PageOfficeCtrl1').InsertDocumentFromURL('/api/ExaminationPaper/Openfile?id='+ this.idArr[i]);
					 this.pNum++;
				}
			}
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window
			window.Create = this.Create;
	    }
	}
</script>

