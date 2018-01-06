
;|AutoLISP例程：读取excel文件   好多讲述用vlisp读取excel文件的文章中都提到
vlax-import-type-library这个函数，其实没有必要。
该函数仅仅是给每个excel对象模型中的属性、方法、对象等引入一个别名，
占内存很大，没有意义。
用vlisp操作excel文件只要了解excel对象模型和
vlax-get-or-create-object 、vlax-get-property、vlax-invoke-method、
vlax-put-property、vlax-safearray-type几个函数就可以了。
下面的例子定义函数(GetCellValueAsList EXCEL_filename sheetName RangeStr)
读取如图所示的excel表格内容，返回list类型。  |;
(defun c:test  ()
(setq	EXCEL_FILELST
	 (dos_getfilem
	   "请选择包含平面和高程精度检测信息汇总的EXCEL格式文件"
	   (if (= filepath nil)
	     (getvar "TEMPPREFIX")
	     filepath)
	   "测量精度检测汇总文件 (*.XLS)|*.XLS|测量精度检测汇总文件 (*.XLSX)|*.XLSX||"))
  (setq path (nth 0 EXCEL_FILELST))
  (setq EXCEL_filename (strcat path (nth 1 EXCEL_FILELST)))
  (setq
    retV (GetCellValueAsList EXCEL_filename "BOM" "A4:E6"))
  (princ))
(defun GetCellValueAsList
       (EXCEL_filename sheetName RangeStr / xl wbs wb shs sh rg cs vvv nms nm ttt)
  (vl-load-com)
  (setq xl (vlax-get-or-create-object "Excel.Application")) ;创建excel程序对象
  (setq wbs (vlax-get-property xl "WorkBooks")) ;获取excel程序对象的工作簿集合对象
  (setq wb (vlax-invoke-method wbs "open" EXCEL_filename));用工作簿集合对象打开指定的excel文件 
  (setq shs (vlax-get-property wb "Sheets")) ;获取刚才打开工作簿的工作表集合
  (setq sh (vlax-get-property shs "Item" sheetName)) ;获取指定的工作表
  (setq rg (vlax-get-property sh "Range" RangeStr)) ;用指定的字符串创建工作表范围对象
  (setq vvv (vlax-get-property rg 'Value)) ;获取范围对象的值
  (setq ttt (vlax-safearray->list (vlax-variant-value vvv))) ;转换为list
  (vlax-invoke-method wb "Close") ;关闭工作簿
  (vlax-invoke-method xl "Quit") ;推出excel对象
  (vlax-release-object xl) ;释放excel对象
  (setq ret ttt))
  (defun c:cs  ()
  (setq	EXCEL_FILELST
	 (dos_getfilem
	   "请选择包含平面和高程精度检测信息汇总的EXCEL格式文件"
	   (if (= filepath nil)
	     (getvar "TEMPPREFIX")
	     filepath)
	   "测量精度检测汇总文件 (*.XLS)|*.XLS|测量精度检测汇总文件 (*.XLSX)|*.XLSX||"))
  (setq path (nth 0 EXCEL_FILELST))
  (setq EXCEL_filename (strcat path (nth 1 EXCEL_FILELST)))
  (vlxls-app-open EXCEL_filename T))
 ;|
Examples:
(setq *xlapp* (vlxls-app-open "C:/test.XLS" T))  ==>  #<VLA-OBJECT _Application 001efd2c>
|;
(defun vlxls-app-open  (XLSFile UnHide / ExcelApp WorkSheet Sheets ActiveSheet Rtn)
  (setq XLSFile (strcase XLSFile))
  (if (null (wcmatch XLSFile "*.XLS"))
    (setq XLSFile (strcat XLSFile ".XLS")))
  (if
    (and (findfile XLSFile) (setq Rtn (vlax-get-or-create-object "Excel.Application")))
     (progn (vlax-invoke-method (vlax-get-property Rtn 'WorkBooks) 'Open XLSFile)
	    (if	UnHide
	      (vla-put-visible Rtn 1)
	      (vla-put-visible Rtn 0))))
  Rtn)