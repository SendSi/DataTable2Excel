### DataTable2Excel
FairyGUI处理多国语言.翻译互导
##### 解决:多国语言时,fairyGUI只能导出xml(简中),此时想搞english,总不能拿着xml给策划去翻译吧
    fairyGUI导出(简中.xml)-->用此工具导出excel(简中.xlsx)-->给策划翻译(english.xlsx)-->用此工具导出xml(english.xml)-->导入到fairyGUI工具中去

##### 1.使用vs2022开发的,,,其他版本未测过
##### 2.调试用Debug,,,,x86的吧 
DataTable2Excel\DataTable2Excel\bin\x86\Debug      exe在这个目录  策划直接用这个目录



### 注意点
    1.xml-->excel 时 要把目标的xlsx删除掉(若存在,会提示删除的)    工具无法覆盖追加  使用时_对应的excel别用wps或office打开哦     
    2.翻译表的追加对比,策划可以 用excel公式进行对比
    eg::: cn.xlsx (11月)   cn.xlsx(12月)   对比出 哪些已翻译过
  