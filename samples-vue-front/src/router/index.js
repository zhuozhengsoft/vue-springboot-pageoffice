import Vue from 'vue'
import VueRouter from 'vue-router'
import Index from '../views/Index.vue'

Vue.use(VueRouter)

const routes = [
  {
    path: '/',
    name: 'Index',
    component: Index
  },
  
  //PageOfficeV5.0新特性
  {
    path: '/CommentOnly/Word',
    component:()=>import('../views/CommentOnly/Word.vue')
  },
  {
    path: '/OnWordSelectionChange/Word',
    component:()=>import("../views/OnWordSelectionChange/Word.vue")
  },
  {
    path: '/Token/index',
    component:()=>import("../views/Token/index.vue")
  },
  
  //基础功能
  {
    path: '/SimpleWord/Word',
    name: 'Word',
    component: () => import('../views/SimpleWord/Word')
  },
  {
    path: '/SimpleWord/Word1',
    component:()=>import("../views/SimpleWord/Word1.vue")
  },
  {
    path: '/SimpleExcel/Excel',
    component:()=>import('../views/SimpleExcel/Excel.vue')
  },
  {
    path: '/SimplePPT/PPT',
    component:()=>import("../views/SimplePPT/PPT.vue")
  },
  {
    path: '/TitleText/Word',
    component:()=>import('../views/TitleText/Word.vue')
  },
  {
    path: '/ControlBars/Word',
    component:()=>import("../views/ControlBars/Word.vue")
  },
  {
    path: '/SetTheme/Word',
    component:()=>import("../views/SetTheme/Word.vue")
  },
  {
    path: '/OpenWord/Word',
    component:()=>import("../views/OpenWord/Word.vue")
  },
  {
    path: '/SaveReturnValue/Word',
    component:()=>import("../views/SaveReturnValue/Word.vue")
  },
  {
    path: '/SendParameters/Word',
    component:()=>import("../views/SendParameters/Word.vue")
  },
  {
    path: '/DataRegionFill/Word',
    component:()=>import("../views/DataRegionFill/Word.vue")
  },
  {
    path: '/ExcelFill/Excel',
    component:()=>import("../views/ExcelFill/Excel.vue")
  },
  {
    path: '/SubmitWord/Word',
    component:()=>import("../views/SubmitWord/Word.vue")
  },
  {
    path: '/SubmitExcel/Excel',
    component:()=>import("../views/SubmitExcel/Excel.vue")
  },
  {
    path: '/InsertSeal/index',
    component:()=>import("../views/InsertSeal/index.vue")
  },
  {
    path: '/InsertSeal/Word/AddSeal/Word1',
    component:()=>import("../views/InsertSeal/Word/AddSeal/Word1.vue")
  },
  {
    path: '/InsertSeal/Word/AddSeal/Word2',
    component:()=>import("../views/InsertSeal/Word/AddSeal/Word2.vue")
  },
  {
    path: '/InsertSeal/Word/AddSeal/Word3',
    component:()=>import("../views/InsertSeal/Word/AddSeal/Word3.vue")
  },
  {
    path: '/InsertSeal/Word/AddSeal/Word4',
    component:()=>import("../views/InsertSeal/Word/AddSeal/Word4.vue")
  },
  {
    path: '/InsertSeal/Word/AddSeal/Word5',
    component:()=>import("../views/InsertSeal/Word/AddSeal/Word5.vue")
  },
  {
    path: '/InsertSeal/Word/AddSeal/Word6',
    component:()=>import("../views/InsertSeal/Word/AddSeal/Word6.vue")
  },
  {
    path: '/InsertSeal/Word/AddSeal/Word7',
    component:()=>import("../views/InsertSeal/Word/AddSeal/Word7.vue")
  },
  {
    path: '/InsertSeal/Word/AddSeal/Word8',
    component:()=>import("../views/InsertSeal/Word/AddSeal/Word8.vue")
  },
  {
    path: '/InsertSeal/Word/BatchAddSeal/index',
    component:()=>import("../views/InsertSeal/Word/BatchAddSeal/index.vue")
  },
  {
    path: '/InsertSeal/Word/AddSeal/Word10',
    component:()=>import("../views/InsertSeal/Word/AddSeal/Word10.vue")
  },
  {
    path: '/InsertSeal/Word/AddSign/Word1',
    component:()=>import("../views/InsertSeal/Word/AddSign/Word1.vue")
  },
  {
    path: '/InsertSeal/Word/AddSign/Word2',
    component:()=>import("../views/InsertSeal/Word/AddSign/Word2.vue")
  },
  {
    path: '/InsertSeal/Word/AddSign/Word3',
    component:()=>import("../views/InsertSeal/Word/AddSign/Word3.vue")
  },
  {
    path: '/InsertSeal/Word/AddSign/Word4',
    component:()=>import("../views/InsertSeal/Word/AddSign/Word4.vue")
  },
  {
    path: '/InsertSeal/Word/AddSign/Word5',
    component:()=>import("../views/InsertSeal/Word/AddSign/Word5.vue")
  },
  {
    path: '/InsertSeal/Excel/AddSeal/Excel1',
    component:()=>import("../views/InsertSeal/Excel/AddSeal/Excel1.vue")
  },
  {
    path: '/InsertSeal/Excel/AddSeal/Excel2',
    component:()=>import("../views/InsertSeal/Excel/AddSeal/Excel2.vue")
  },
  {
    path: '/InsertSeal/Excel/AddSeal/Excel3',
    component:()=>import("../views/InsertSeal/Excel/AddSeal/Excel3.vue")
  },
  {
    path: '/InsertSeal/Excel/AddSeal/Excel4',
    component:()=>import("../views/InsertSeal/Excel/AddSeal/Excel4.vue")
  },
  {
    path: '/InsertSeal/Excel/AddSeal/Excel5',
    component:()=>import("../views/InsertSeal/Excel/AddSeal/Excel5.vue")
  },
  {
    path: '/InsertSeal/Excel/AddSign/Excel1',
    component:()=>import("../views/InsertSeal/Excel/AddSign/Excel1.vue")
  },
  {
    path: '/InsertSeal/Excel/AddSign/Excel2',
    component:()=>import("../views/InsertSeal/Excel/AddSign/Excel2.vue")
  },
  {
    path: '/InsertSeal/Excel/AddSign/Excel3',
    component:()=>import("../views/InsertSeal/Excel/AddSign/Excel3.vue")
  },
  {
    path: '/InsertSeal/PDF/AddSeal/PDF1',
    component:()=>import("../views/InsertSeal/PDF/AddSeal/PDF1.vue")
  },
  {
    path: '/InsertSeal/PDF/AddSign/PDF1',
    component:()=>import("../views/InsertSeal/PDF/AddSign/PDF1.vue")
  },
  
  
  {
    path: '/CommandCtrl/Word',
    component:()=>import("../views/CommandCtrl/Word.vue")
  },
  {
    path: '/WordSetTable/Word',
    component:()=>import("../views/WordSetTable/Word.vue")
  },
  {
    path: '/WordDataTag2/Word',
    component:()=>import("../views/WordDataTag2/Word.vue")
  },
  {
    path: '/CustomToolButton/Word',
    component:()=>import("../views/CustomToolButton/Word.vue")
  },
  {
    path: '/AfterDocOpened/Word',
    component:()=>import("../views/AfterDocOpened/Word.vue")
  },
  {
    path: '/JsControlBars/Word',
    component:()=>import("../views/JsControlBars/Word.vue")
  },
  {
    path: '/ConcurrencyCtrl/index',
    component:()=>import("../views/ConcurrencyCtrl/index.vue")
  },
  {
    path: '/ConcurrencyCtrl/Word',
    component:()=>import("../views/ConcurrencyCtrl/Word.vue")
  },
  {
    path: '/ExcelTable/Excel',
    component:()=>import("../views/ExcelTable/Excel.vue")
  },
  {
    path: '/SaveAsHTML/Word',
    component:()=>import("../views/SaveAsHTML/Word.vue")
  },
  {
    path: '/SaveAsMHT/Word',
    component:()=>import("../views/SaveAsMHT/Word.vue")
  },
  {
    path: '/BeforeAndAfterSave/Word',
    component:()=>import("../views/BeforeAndAfterSave/Word.vue")
  },
  {
    path: '/POBrowser2/Word',
    component:()=>import("../views/POBrowser2/Word.vue")
  },
  {
    path: '/SaveDataAndFile/Word',
    component:()=>import("../views/SaveDataAndFile/Word.vue")
  },
  {
    path: '/ImportWordData/Word',
    component:()=>import("../views/ImportWordData/Word.vue")
  },
  {
    path: '/ImportExcelData/Excel',
    component:()=>import("../views/ImportExcelData/Excel.vue")
  },
  {
    path: '/WordDisableRight/Word',
    component:()=>import("../views/WordDisableRight/Word.vue")
  },
  {
    path: '/ExcelDisableRight/Excel',
    component:()=>import("../views/ExcelDisableRight/Excel.vue")
  },
  {
    path: '/POBrowserTopic/index',
    component:()=>import("../views/POBrowserTopic/index.vue")
  },
  {
    path: '/POBrowserTopic/Word1',
    component:()=>import("../views/POBrowserTopic/Word1.vue")
  },
  {
    path: '/POBrowserTopic/Word2',
    component:()=>import("../views/POBrowserTopic/Word2.vue")
  },
  {
    path: '/NoFrame/Word',
    component:()=>import("../views/NoFrame/Word.vue")
  },
  {
    path: '/RevisionOnly/Word',
    component:()=>import("../views/RevisionOnly/Word.vue")
  },
  
  //高级功能
  {
    path: '/ReadOnly/Word',
    component:()=>import("../views/ReadOnly/Word.vue")
  },
  {
    path: '/DataBase/Word',
    component:()=>import("../views/DataBase/Word.vue")
  },
  {
    path: '/CreateWord/index',
    component:()=>import("../views/CreateWord/index.vue")
  },
  {
    path: '/CreateWord/createWord',
    component:()=>import("../views/CreateWord/createWord.vue")
  },
  {
    path: '/CreateWord/Word',
    component:()=>import("../views/CreateWord/Word.vue")
  },
  {
    path: '/POPDF/PDF',
    component:()=>import("../views/POPDF/PDF.vue")
  },
  {
    path: '/SaveAsPDF/Word',
    component:()=>import("../views/SaveAsPDF/Word.vue")
  },
  {
    path: '/WordResWord/Word',
    component:()=>import("../views/WordResWord/Word.vue")
  },
  {
    path: '/WordResImage/Word',
    component:()=>import("../views/WordResImage/Word.vue")
  },
  {
    path: '/WordResExcel/Word',
    component:()=>import("../views/WordResExcel/Word.vue")
  },
  {
    path: '/AddWaterMark/Word',
    component:()=>import("../views/AddWaterMark/Word.vue")
  },
  {
    path: '/WordDataTag/Word',
    component:()=>import("../views/WordDataTag/Word.vue")
  },
  {
    path: '/DataRegionCreate/Word',
    component:()=>import("../views/DataRegionCreate/Word.vue")
  },
  {
    path: '/RunMacro/Word',
    component:()=>import("../views/RunMacro/Word.vue")
  },
  {
    path: '/FileMakerSingle/index',
    component:()=>import("../views/FileMakerSingle/index.vue")
  },
  {
    path: '/FileMakerSingle/Word',
    component:()=>import("../views/FileMakerSingle/Word.vue")
  },
  {
    path: '/FileMakerSingle/OpenWord',
    component:()=>import("../views/FileMakerSingle/OpenWord.vue")
  },
  {
    path: '/WordTable/Word',
    component:()=>import("../views/WordTable/Word.vue")
  },
  {
    path: '/WordHandDraw/Word',
    component:()=>import("../views/WordHandDraw/Word.vue")
  },
  {
    path: '/DataRegionTable/Word',
    component:()=>import("../views/DataRegionTable/Word.vue")
  },
  {
    path: '/DataRegionText/Word',
    component:()=>import("../views/DataRegionText/Word.vue")
  },
  {
    path: '/SetDrByUserWord/index',
    component:()=>import("../views/SetDrByUserWord/index.vue")
  },
  {
    path: '/SetDrByUserWord2/index',
    component:()=>import("../views/SetDrByUserWord2/index.vue")
  },
  {
    path: '/SetHandDrawByUser/index',
    component:()=>import("../views/SetHandDrawByUser/index.vue")
  },
  {
    path: '/MergeWordCell/Word',
    component:()=>import("../views/MergeWordCell/Word.vue")
  },
  {
    path: '/ClickDataRegion/Word',
    component:()=>import("../views/ClickDataRegion/Word.vue"),
  },
  {
    path: '/ClickDataRegion/HTMLPage',
    component:()=>import("../views/ClickDataRegion/HTMLPage.vue"),
  },
  {
    path: '/MergeExcelCell/Excel',
    component:()=>import("../views/MergeExcelCell/Excel.vue")
  },
  {
    path: '/SetXlsTableByUser/index',
    component:()=>import("../views/SetXlsTableByUser/index.vue")
  },
  {
    path: '/SetExcelCellBorder/Excel',
    component:()=>import("../views/SetExcelCellBorder/Excel.vue")
  },
  {
    path: '/SetExcelCellText/Excel',
    component:()=>import("../views/SetExcelCellText/Excel.vue")
  },
  {
    path: '/DataRegionFill2/Word',
    component:()=>import("../views/DataRegionFill2/Word.vue")
  },
  {
    path: '/ExcelCellClick/Excel',
    component:()=>import("../views/ExcelCellClick/Excel.vue")
  },
  {
    path: '/ExcelCellClick/Select',
    component:()=>import("../views/ExcelCellClick/Select1.vue")
  },
  {
    path: '/ExcelFill2/Excel',
    component:()=>import("../views/ExcelFill2/Excel.vue")
  },
  {
    path: '/DataRegionEdit/Word',
    component:()=>import("../views/DataRegionEdit/Word.vue")
  },
  {
    path: '/DataRegionEdit/DataRegionDlg',
    component:()=>import("../views/DataRegionEdit/DataRegionDlg.vue")
  },
  {
    path: '/DataTagEdit/Word',
    component:()=>import("../views/DataTagEdit/Word.vue")
  },
  {
    path: '/DataTagEdit/DataTagDlg',
    component:()=>import("../views/DataTagEdit/DataTagDlg.vue")
  },
  {
    path: '/DefinedNameCell/Excel',
    component:()=>import("../views/DefinedNameCell/Excel.vue")
  },
  {
    path: '/DefinedNameTable/index',
    component:()=>import("../views/DefinedNameTable/index.vue")
  },
  {
    path: '/DefinedNameTable/Excel',
    component:()=>import("../views/DefinedNameTable/Excel.vue")
  },
  {
    path: '/DefinedNameTable/Excel2',
    component:()=>import("../views/DefinedNameTable/Excel2.vue")
  },
  {
    path: '/DefinedNameTable/Excel4',
    component:()=>import("../views/DefinedNameTable/Excel4.vue")
  },
  {
    path: '/DefinedNameTable/Excel5',
    component:()=>import("../views/DefinedNameTable/Excel5.vue")
  },
  {
    path: '/DefinedNameTable/Excel6',
    component:()=>import("../views/DefinedNameTable/Excel6.vue")
  },
  {
    path: '/FileMakerPDF/index',
    component:()=>import("../views/FileMakerPDF/index.vue")
  },
  {
    path: '/FileMakerPDF/OpenPDF',
    component:()=>import("../views/FileMakerPDF/OpenPDF.vue")
  },
  {
    path: '/FileMakerPDF/PDF',
    component:()=>import("../views/FileMakerPDF/PDF.vue")
  },
  {
    path: '/WordCompare/Word',
    component:()=>import("../views/WordCompare/Word.vue")
  },
  {
    path: '/WordTextBox/Word',
    component:()=>import("../views/WordTextBox/Word.vue")
  },
  {
    path: '/WordRibbonCtrl/Word',
    component:()=>import("../views/WordRibbonCtrl/Word.vue")
  },
  {
    path: '/ExcelRibbonCtrl/Excel',
    component:()=>import("../views/ExcelRibbonCtrl/Excel.vue")
  },
  {
    path: '/SplitWord/Word',
    component:()=>import("../views/SplitWord/Word.vue")
  },
  {
    path: '/CommentsList/Word',
    component:()=>import("../views/CommentsList/Word.vue")
  },
  {
    path: '/RevisionsList/Word',
    component:()=>import("../views/RevisionsList/Word.vue")
  },
  {
    path: '/HandDrawsList/Word',
    component:()=>import("../views/HandDrawsList/Word.vue")
  },
  {
    path: '/WordCreateTable/Word',
    component:()=>import("../views/WordCreateTable/Word.vue")
  },
  {
    path: '/RunMacro2/Word',
    component:()=>import("../views/RunMacro2/Word.vue")
  },
  {
    path: '/PDFSearch/PDF',
    component:()=>import("../views/PDFSearch/PDF.vue")
  },
  {
    path: '/SaveFirstPageAsImg/Word',
    component:()=>import("../views/SaveFirstPageAsImg/Word.vue")
  },
  {
    path: '/ExcelAdjustRC/Excel',
    component:()=>import("../views/ExcelAdjustRC/Excel.vue")
  },
  {
    path: '/WordDeleteRow/Word',
    component:()=>import("../views/WordDeleteRow/Word.vue")
  },
  {
    path: '/InsertPageBreak2/Word',
    component:()=>import("../views/InsertPageBreak2/Word.vue")
  },
  {
    path: '/ExcelInsertImage/Excel',
    component:()=>import("../views/ExcelInsertImage/Excel.vue")
  },
  {
    path: '/WordTableSetImg/Word',
    component:()=>import("../views/WordTableSetImg/Word.vue")
  },
  {
    path: '/WordTableBorder/Word',
    component:()=>import("../views/WordTableBorder/Word.vue")
  },
  {
    path: '/ExtractImage/Word',
    component:()=>import("../views/ExtractImage/Word.vue")
  },
  {
    path: '/OpenImage/Image',
    component:()=>import("../views/OpenImage/Image.vue")
  },
  {
    path: '/DisableCopyOut/Word',
    component:()=>import("../views/DisableCopyOut/Word.vue")
  },
  {
    path: '/InsertImageSetSize/Word',
    component:()=>import("../views/InsertImageSetSize/Word.vue")
  },
  {
    path: '/RunMacroForDocm/Word',
    component:()=>import("../views/RunMacroForDocm/Word.vue")
  },
  
  //综合演示
  {
    path: '/FileMaker/Default',
    component:()=>import("../views/FileMaker/Default.vue")
  },
  {
    path: '/ExaminationPaper/index',
    component:()=>import("../views/ExaminationPaper/index.vue")
  },
  {
    path: '/ExaminationPaper/Word',
    component:()=>import("../views/ExaminationPaper/Word.vue")
  },
  {
    path: '/ExaminationPaper/Compose',
    component:()=>import("../views/ExaminationPaper/Compose.vue")
  },
  {
    path: '/ExaminationPaper/Compose2',
    component:()=>import("../views/ExaminationPaper/Compose2.vue")
  },
  {
    path: '/WordParagraph/Word',
    component:()=>import("../views/WordParagraph/Word.vue")
  },
  {
    path: '/DrawExcel/Excel',
    component:()=>import("../views/DrawExcel/Excel.vue")
  },
  {
    path: '/TaoHong/index',
    component:()=>import("../views/TaoHong/index.vue")
  },
  {
    path: '/TaoHong/Word',
    component:()=>import("../views/TaoHong/Word.vue")
  },
  {
    path: '/TaoHong/taoHong',
    component:()=>import("../views/TaoHong/taoHong.vue")
  },
  {
    path: '/TaoHong/readOnly',
    component:()=>import("../views/TaoHong/readOnly.vue")
  },
  {
    path: '/WordSalaryBill/index',
    component:()=>import("../views/WordSalaryBill/index.vue")
  },
  {
    path: '/WordSalaryBill/Word',
    component:()=>import("../views/WordSalaryBill/Word.vue")
  },
  {
    path: '/WordSalaryBill/OpenFile',
    component:()=>import("../views/WordSalaryBill/OpenFile.vue")
  },
  {
    path: '/WordSalaryBill/Compose',
    component:()=>import("../views/WordSalaryBill/Compose.vue")
  },
  {
    path: '/PrintFiles/Default',
    component:()=>import("../views/PrintFiles/Default.vue")
  },
  {
    path: '/SaveAndSearch/index',
    component:()=>import("../views/SaveAndSearch/index.vue")
  },
  {
    path: '/SaveAndSearch/Word',
    component:()=>import("../views/SaveAndSearch/Word.vue")
  },
  {
    path: '/FileMakerConvertPDFs/index',
    component:()=>import("../views/FileMakerConvertPDFs/index.vue")
  },
  
  
  
  //其他技巧
  {
    path: '/DeleteRow/Word',
    component:()=>import("../views/DeleteRow/Word.vue")
  },
  {
    path: '/HiddenRulars/Word',
    component:()=>import("../views/HiddenRulars/Word.vue")
  },
  {
    path: '/WordAddBKMK/Word',
    component:()=>import("../views/WordAddBKMK/Word.vue")
  },
  {
    path: '/WordLocateBKMK/Word',
    component:()=>import("../views/WordLocateBKMK/Word.vue")
  },
  {
    path: '/WordHyperLink/Word',
    component:()=>import("../views/WordHyperLink/Word.vue")
  },
  {
    path: '/WordMergeCell/Word',
    component:()=>import("../views/WordMergeCell/Word.vue")
  },
  {
    path: '/WordGetSelection/Word',
    component:()=>import("../views/WordGetSelection/Word.vue")
  },
  {
    path: '/WordGoToPage/Word',
    component:()=>import("../views/WordGoToPage/Word.vue")
  },
  {
    path: '/JsOpXlsCellText/Excel',
    component:()=>import("../views/JsOpXlsCellText/Excel.vue")
  },
  {
    path: '/InsertPageBreak/Word',
    component:()=>import("../views/InsertPageBreak/Word.vue")
  },
  {
    path: '/WordDelBKMK/Word',
    component:()=>import("../views/WordDelBKMK/Word.vue")
  },
  {
    path: '/InsertImgForJs/Word',
    component:()=>import("../views/InsertImgForJs/Word.vue")
  },
  {
    path: '/InsertImgWaterMark/Word',
    component:()=>import("../views/InsertImgWaterMark/Word.vue")
  },
  
  //PAGEOFFICE V4.0新特性
  {
    path: '/POBrowser/Word',
    component:()=>import("../views/POBrowser/Word.vue")
  },
  {
    path: '/HtmlWord/Word',
    component:()=>import("../views/HtmlWord/Word.vue")
  },
  {
    path: '/CallParentFunction/index',
    component:()=>import("../views/CallParentFunction/index.vue")
  },
  {
    path: '/CallParentFunction/Word',
    component:()=>import("../views/CallParentFunction/Word.vue")
  },
  
  {
    path: '/PromptSave/Word',
    component:()=>import("../views/PromptSave/Word.vue")
  },
  {
    path: '/OpenWindowModeless/Word',
    component:()=>import("../views/OpenWindowModeless/Word.vue")
  },
  {
    path: '/GetParentParamValue/Word',
    component:()=>import("../views/GetParentParamValue/Word.vue")
  },
  {
    path: '/OpenWindow/Word',
    component:()=>import("../views/OpenWindow/Word.vue")
  },
  
  //仅支持IE浏览器的PAGEOFFICE集成方式
  {
    path: '/SimpleWord2/Word',
    component:()=>import('../views/SimpleWord2/Word.vue')
  },
  {
    path: '/SimpleExcel2/Excel',
    component:()=>import('../views/SimpleExcel2/Excel.vue')
  },
  {
    path: '/POPDF2/PDF',
    component:()=>import('../views/POPDF2/PDF.vue')
  }
]

const router = new VueRouter({
	mode:'history',
	routes
})

export default router
