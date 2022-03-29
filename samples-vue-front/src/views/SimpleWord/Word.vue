<template>
  <div class="Word">
    <div style="height: 800px; width: auto" v-html="poHtmlCode" />
  </div>
</template>
<script>
const axios = require("axios");
export default {
  name: "Word",
  data() {
    return {
      message: "SimpleWord!",
      poHtmlCode: "",
    };
  },
  created: function () {
    //由于vue中的axios拦截器给请求加token都得是ajax请求，所以这里建议axios方式去请求后台打开文件的controller方法
    axios
      .post("/api/SimpleWord/Word")
      .then((response) => {
        this.poHtmlCode = response.data;
      })
      .catch(function (err) {
        console.log(err);
      });
  },
  methods: {
    //控件中的一些常用方法都在这里调用，比如保存，打印等等
    /**
     * Save()，callParent()等方法都是/api/SimpleWord/Word这个后台controller中PageOfficeCtrl控件通过poCtrl.addCustomToolButton定义的方法。
     */
    Save() {
      document.getElementById("PageOfficeCtrl1").WebSave();
    },
    SaveAs() {
      document.getElementById("PageOfficeCtrl1").ShowDialog(3);
    },
    PrintSet() {
      document.getElementById("PageOfficeCtrl1").ShowDialog(5);
    },
    PrintFile() {
      document.getElementById("PageOfficeCtrl1").ShowDialog(4);
    },
    Close() {
      window.external.close();
    },
    IsFullScreen() {
      document.getElementById("PageOfficeCtrl1").FullScreen =
        !document.getElementById("PageOfficeCtrl1").FullScreen;
    },
  },
  mounted: function () {
    // 将PageOffice控件中的方法通过mounted挂载到window对象上，只有挂载后才能被vue组件识别
    window.Save = this.Save;
    window.SaveAs = this.SaveAs;
    window.PrintSet = this.PrintSet;
    window.PrintFile = this.PrintFile;
    window.Close = this.Close;
    window.IsFullScreen = this.IsFullScreen;
  },
};
</script>

