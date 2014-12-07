# coding: utf-8
require 'spreadsheet'
require 'yaml'
Spreadsheet.client_encoding = 'UTF-8'

#读取配置文件
options = YAML.load(File.open("option.yml"))
list = options["list"]
template = options["template"]
map = options["map"]
fname_col = options["fname"]

#打开工作表
workbook = Spreadsheet.open list
templatebook = Spreadsheet.open template
@template = templatebook.worksheet 0
worksheet = workbook.worksheet 0

#定义函数：将data数据写入模板template.range(cell)单元格中, cell表示形式如"A1"
def mapTemplate (cell, data)

  #将cell中用字母表式的列转换为数字表示方式，如A -> 0
  cell_col_str = /^[a-z]+/.match(cell.downcase)[0]
  sum = 0
  len = cell_col_str.size
  cell_col_str.each_byte do |c|
    sum = (c - 96 ) * 26 ** (len -1) + sum
    len = len-1
  end 
  cell_col = sum - 1
  
  cell_row = /\d+/.match(cell)[0].to_i - 1

  #写入模板
  @template.rows[cell_row][cell_col]= data
end

#遍历工作表
worksheet.each do |row|
  
  # 行首为数字序号的行执行映射，其它的忽略。
  if row[0].kind_of? Integer
    map.each do |k,v|
      cells = Array.new
      cells = v.split(' ')
      cells.each do |cell|
        mapTemplate(cell, row[k])
      end
    end

    # 另存文件，文件名由工作表中fname_col列字段确定。
    fname = row[fname_col].to_s + ".xls"
    templatebook.write fname
  end
end

