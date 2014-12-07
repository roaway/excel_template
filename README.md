# 一个很简单的Ruby脚本
利用了Spreadsheet模块处理Excel文件，YAML模块读取脚本配置。实现一个简单功能：将一张条目汇总表list.xls中数据映射到另一张模板表template.xls,生成每个条目以template.xls为模板的Excel文件。举例而言，现有一张人员信息的汇总表和个人简历模板表，通过设置映射关系，可生成汇总表中每个人员的个人简历表。

# 配置文件：option.yml
<pre>
list: list.xls            #汇总表
template: template.xls    #模板文件
map:                      #映射关系
  1: B2 B17               #将汇总表的第"1"列映射到模板文件的"B2" "B17"单元格中，多个单元格时用空格分开
  2: D2 D17
fname: 1                  #以汇总表第“1”列作为文件名命名。
</pre>
