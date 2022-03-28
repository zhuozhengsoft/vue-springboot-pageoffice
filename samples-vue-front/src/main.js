// The Vue build version to load with the `import` command
// (runtime-only or standalone) has been set in webpack.base.conf with an alias.
import Vue from 'vue'
import App from './App'
import router from './router'
import  VueResource  from 'vue-resource'
import axios from  'axios'


Vue.prototype.HOST = '/api'
Vue.config.productionTip = false
Vue.use(VueResource);

// 添加请求拦截器

axios.interceptors.request.use(function (config) {
  let token="token1";
  // 在发送请求之前做些什么
  // 将每个页面header添加token
    config.headers.common['Authorization'] =token;
  return config
}, function (error) {
  //给所有请求都加上token
  router.push('/!*')
  return Promise.reject(error)
})
// 添加响应拦截器
axios.interceptors.response.use(function (response) {
  // 对响应数据做点什么
  return response
}, function (error) {
  // 对响应错误做点什么
  if (error.response) {
    switch (error.response.status) {
      case 401:
        store.commit('del_token')
        router.push('/login')
    }
  }
  return Promise.reject(error)
})


/* eslint-disable no-new */
new Vue({
  el: '#app',
  router,
  components: { App },
  template: '<App/>'
})
