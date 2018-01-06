
;|AutoLISP���̣���ȡexcel�ļ�   �öི����vlisp��ȡexcel�ļ��������ж��ᵽ
vlax-import-type-library�����������ʵû�б�Ҫ��
�ú��������Ǹ�ÿ��excel����ģ���е����ԡ����������������һ��������
ռ�ڴ�ܴ�û�����塣
��vlisp����excel�ļ�ֻҪ�˽�excel����ģ�ͺ�
vlax-get-or-create-object ��vlax-get-property��vlax-invoke-method��
vlax-put-property��vlax-safearray-type���������Ϳ����ˡ�
��������Ӷ��庯��(GetCellValueAsList EXCEL_filename sheetName RangeStr)
��ȡ��ͼ��ʾ��excel������ݣ�����list���͡�  |;
(defun c:test  ()
(setq	EXCEL_FILELST
	 (dos_getfilem
	   "��ѡ�����ƽ��͸߳̾��ȼ����Ϣ���ܵ�EXCEL��ʽ�ļ�"
	   (if (= filepath nil)
	     (getvar "TEMPPREFIX")
	     filepath)
	   "�������ȼ������ļ� (*.XLS)|*.XLS|�������ȼ������ļ� (*.XLSX)|*.XLSX||"))
  (setq path (nth 0 EXCEL_FILELST))
  (setq EXCEL_filename (strcat path (nth 1 EXCEL_FILELST)))
  (setq
    retV (GetCellValueAsList EXCEL_filename "BOM" "A4:E6"))
  (princ))
(defun GetCellValueAsList
       (EXCEL_filename sheetName RangeStr / xl wbs wb shs sh rg cs vvv nms nm ttt)
  (vl-load-com)
  (setq xl (vlax-get-or-create-object "Excel.Application")) ;����excel�������
  (setq wbs (vlax-get-property xl "WorkBooks")) ;��ȡexcel�������Ĺ��������϶���
  (setq wb (vlax-invoke-method wbs "open" EXCEL_filename));�ù��������϶����ָ����excel�ļ� 
  (setq shs (vlax-get-property wb "Sheets")) ;��ȡ�ղŴ򿪹������Ĺ�������
  (setq sh (vlax-get-property shs "Item" sheetName)) ;��ȡָ���Ĺ�����
  (setq rg (vlax-get-property sh "Range" RangeStr)) ;��ָ�����ַ�������������Χ����
  (setq vvv (vlax-get-property rg 'Value)) ;��ȡ��Χ�����ֵ
  (setq ttt (vlax-safearray->list (vlax-variant-value vvv))) ;ת��Ϊlist
  (vlax-invoke-method wb "Close") ;�رչ�����
  (vlax-invoke-method xl "Quit") ;�Ƴ�excel����
  (vlax-release-object xl) ;�ͷ�excel����
  (setq ret ttt))
  (defun c:cs  ()
  (setq	EXCEL_FILELST
	 (dos_getfilem
	   "��ѡ�����ƽ��͸߳̾��ȼ����Ϣ���ܵ�EXCEL��ʽ�ļ�"
	   (if (= filepath nil)
	     (getvar "TEMPPREFIX")
	     filepath)
	   "�������ȼ������ļ� (*.XLS)|*.XLS|�������ȼ������ļ� (*.XLSX)|*.XLSX||"))
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