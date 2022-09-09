module.exports = {
    devServer: {
        proxy: {
        '/api': {
          target: 'http://localhost:8081/samples-springboot-back', //"/api"对应后端项目"http://localhost:8081/samples-springboot-back"地址 
          ws: true,
          changeOrigin: true, // 允许跨域
          pathRewrite: {
           '^/api': ''   // 标识替换，使用 '/api' 代替真实的接口地址
          }
        }
      }
    },
	publicPath:"/",
	
	//node_modules里的依赖默认是不会编译的,会导致es6语法在ie中的语法报错,根据报错找到对应的文件夹指定该文件夹或文件需要编译.
	transpileDependencies: ["sockjs-client"]
  }