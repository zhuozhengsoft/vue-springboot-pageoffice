<template>
	<div class="Word">
		<div class="zz-content mc clearfix pd-28" align="center">
		
		    <div class="demo mc">
		        <h2 class="fs-16">PageOffice 实现Word文档的在线编辑保存和全文关键字搜索</h2>
		    </div>
		
		    <div class="demo mc">
		        <h3 class="fs-12">搜索文件</h3>
		        <!-- <form id="form1" action="index" method="get"> -->
				<form id="form1">
		            <table class="text" cellspacing="0" cellpadding="0" border="0">
		                <tr>
		                    <td style="font-size: 9pt" align="left">
		                        通过文档内容中的关键字搜索文档&nbsp;&nbsp;&nbsp;
		                    </td>
		                    <td align="center">
		                        <input name="Input_KeyWord" id="Input_KeyWord" type="text" @focus="getFocus()"
		                               @blur="lostFocus()"
		                               class="boder" style="width: 180px;" value="请输入关键字"/>
		                    </td>
		                    <td style="width: 221px;">
		                        &nbsp;
		                        <input type="button" value="搜索" @click="mySubmit();" style=" width:86px;"/>
		                    </td>
		                </tr>
		                <tr>
		                    <td>&nbsp;</td>
		                    <td colspan="2">&nbsp;<span style="color:Gray;">热门搜索：</span>
		                        <a href="#" style="color:#00217d;" @click="copyKeyToInput('网站');">网站</a>
		                        <a href="#" style="color:#00217d;" @click="copyKeyToInput('软件');">软件</a>
		                        <a href="#" style="color:#00217d;" @click="copyKeyToInput('字体');">字体</a></td>
		                </tr>
		            </table>
		        </form>
		    </div>
		
		    <div class="zz-talbeBox mc">
		        <h3 class="fs-12">文档列表</h3>
		
		        <table class="zz-talbe">
		            <thead>
						<tr>
							<th width="90%">文档名称</th>
							<th width="100%">操作</th>
						</tr>
		            </thead>
		
		            <tbody>
						<tr v-for="docSearch in docSearchs" v-model="docSearchs">
							<td>{{docSearch.fileName}}</td>
							<td v-if="value != ''&& value !=null">
								<input type="button" @click="editWord1(docSearch.id)" value="编辑1"></input>
							</td>
							<td v-else>
								<input type="button" @click="editWord2(docSearch.id)" value="编辑2"></input>
							</td>
										
							
						</tr>
		            </tbody>
		
		        </table>
				
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
	        docSearchs: {},
			value:''
	      }
	    },
	    created: function(){
	      //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里必须是axios方式去请求后台打开文件的controller
	      axios.post("/api/SaveAndSearch/index").then((response) => {
	        this.docSearchs = response.data.docSearchs;
	      }).catch(function (err) {
	        console.log(err)
	      })
	    },
	    methods:{
			editWord1(id){
				POBrowser.openWindowModeless('SaveAndSearch/Word?id='+id, 'width=1200px;height=800px;',this.value);
			},
			editWord2(id){
				POBrowser.openWindowModeless('Word?id='+id, 'width=1200px;height=800px;');
			},
	      //控件中的一些常用方法都在这里调用，比如保存，打印等等
			onColor(td) {
				td.style.backgroundColor = '#87ffc7';
			},
	
			offColor(td) {
				td.style.backgroundColor = '';
			},
	
			getFocus() {
				var str = document.getElementById("Input_KeyWord").value;
				if (str == "请输入关键字") {
					document.getElementById("Input_KeyWord").value = "";
				}
	
			},
	
			lostFocus() {
				var str = document.getElementById("Input_KeyWord").value;
				if (str.length <= 0) {
					document.getElementById("Input_KeyWord").value = "请输入关键字";
				}
			},
	
			copyKeyToInput(key) {
				document.getElementById("Input_KeyWord").value = key;
			},
	
			mySubmit() {
				this.value = encodeURI(document.getElementById("Input_KeyWord").value);
				document.getElementById("form1").submit = axios.post("/api/SaveAndSearch/index?Input_KeyWord="+this.value).then((response) => {
															this.docSearchs = response.data.docSearchs;
														  }).catch(function (err) {
															console.log(err)
														  })
												
			}
	    },
	    mounted: function(){
	      // 将vue中的方法赋值给window

	    }
	}
</script>

