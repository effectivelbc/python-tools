# -*- coding: utf-8 -*-
import pymysql
import xlwt
import re
import random


#翻译字典 常见的翻译
chinese_column_name = {
    'create_time'    : '创建时间',
    'update_time'    : '修改时间',
    'modify_time'    : '修改时间',
    'createtime'     : '创建时间',
    'updatetime'     : '修改时间',
    'modifytime'     : '修改时间',
    'create_date'    : '创建时间',
    'update_date'    : '修改时间',
    'modify_date'    : '修改时间',
    'createdate'     : '创建时间',
    'updatedate'     : '修改时间',
    'modifydate'     : '修改时间',
    'state'          : '状态',
    'orders'         : '排序',
    'sort'           : '排序',
    'type'           : '类型',
    'id'             : 'ID',
    'ID'             : 'ID',
    'is_valid'       : '是否有效',
    'is_active'      : '是否有效',
    'name'           : '名称',
    'status'         : '状态',
    'number'         : '数量',
    'creater'        : '创建人',
    'creator'        : '创建人',
    'source'         : '来源',
    'phone'          : '电话',
    'remark'         : '备注',
    'contact'        : '联系人',
    'start_date'     : '开始时间',
    'end_date'       : '结束时间',
    'starttime'      : '开始时间',
    'endtime'        : '结束时间',
    'startdate'      : '开始时间',
    'enddate'        : '结束时间',
    'start_time'     : '开始时间',
    'end_time'       : '结束时间',
    'money'          : '金额',
    'content'        : '内容',
    'date'           : '日期',
    'amount'         : '数量',
    'address'        : '地址',
    'gender'         : '性别',
    'sex'            : '性别',
    'age'            : '年龄',
    'quantity'       : '数量',
    'url'            : 'url',
    'img'            : '图片',
    'image'          : '图片',
    'operator'       : '操作人', 
    'area'           : '区域',
    'contactor'      : '联系人',
    'mobile'         : '手机',
    'introducer'     : '介绍人',
    'logo'           : '图标logo',
    'country'        : '国家',
    'product_name'   : '产品名称',
    'birthday'       : '生日',
    'wechat_account' : '微信账户',
    'balance'        : '余额', 
    'time'           : '时间',
    'star'           : '星级',
    'days'           : '天数',
    'supplier'       : '供应商',
    'store'          : '门店',
    'supplier_id'    : '供应商id',
    'store_id'       : '门店id',
    
}







# 获取字符串长度，一个中文的长度为2
def len_byte(value):
    length = len(value)
    utf8_length = len(value.encode('utf-8'))
    length = (utf8_length - length) / 2 + length
    return int(length)



#字体颜色默认黑色 背景颜色默认为白色 边框颜色默认为黑色
def setStyle(name, height, font_color = 0, pattern_color = 1, border_color = 0, bold = False):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    # 字体类型：比如宋体、仿宋也可以是汉仪瘦金书繁
    font.name = name
    #是否粗体
    font.bold = bold
    # 设置字体颜色
    font.colour_index = font_color
    # 字体大小
    font.height = height
    # 定义格式
    style.font = font
    
    # 定义边框
    # borders.left = xlwt.Borders.THIN
    # NO_LINE： 官方代码中NO_LINE所表示的值为0，没有边框
    # THIN： 官方代码中THIN所表示的值为1，边框为实线
    borders = xlwt.Borders()
    borders.left = border_color
    borders.left = xlwt.Borders.THIN
    borders.right = border_color
    borders.right = xlwt.Borders.THIN
    borders.top = border_color
    borders.top = xlwt.Borders.THIN
    borders.bottom = border_color
    borders.bottom = xlwt.Borders.THIN

    # 定义格式
    style.borders = borders
    
    # 设置背景颜色
    pattern = xlwt.Pattern()
    # 设置背景颜色的模式
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN

    # 背景颜色
    pattern.pattern_fore_colour = pattern_color

    style.pattern = pattern
    
    #设置居中
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER  #水平方向
    alignment.vert = xlwt.Alignment.VERT_CENTER      
    
    style.alignment = alignment
    
    
    return style


file = xlwt.Workbook(encoding='utf-8',style_compression=2)


conn = pymysql.connect(
    host="47.104.68.26",
    user="root",
    passwd="ZXClogan123!",
    port=3306
)

cur = conn.cursor()

cur.execute('SHOW DATABASES')

#数据库的元组
database_all = cur.fetchall()

print(database_all)

#获取库名是hy_dba的
conn.select_db(database_all[3][0])
cur.execute('SHOW TABLES')
ret = cur.fetchall()
print("一共有", len(ret), "个表")

sheet_name_set = []

#获取后100张表的名字
for i in ret[25:]:
    #if i[0] == 'hy_ticket_hotel':
    #新建一个名字为表名的sheet 可重复写
    sheet_name = i[0]
    if len(sheet_name) > 31:
        #截取
        sheet_name = sheet_name[:31]
    if sheet_name not in sheet_name_set:
        sheet_name_set.append(sheet_name)
        table = file.add_sheet(sheet_name, cell_overwrite_ok=True)
    else:
        #有重名
        while sheet_name in sheet_name_set:
            sheet_name = sheet_name[0:-1] + str(random.randint(0,9))
        sheet_name_set.append(sheet_name)
        table = file.add_sheet(sheet_name, cell_overwrite_ok=True)
    #背景蓝色
    table.write(0, 0, '属性英文名称', setStyle('宋体', height=200, pattern_color=7, bold=True))
    table.write(1, 0, '属性中文名称', setStyle('宋体', height=200, pattern_color=7, bold=True))
    table.write(2, 0, '字段类型', setStyle('宋体', height=200, pattern_color=7, bold=True))
    table.write(3, 0, '是否必填', setStyle('宋体', height=200, pattern_color=7, bold=True))
    table.write(4, 0, '备注', setStyle('宋体', height=200, pattern_color=7, bold=True))
    for j in range(5):
        #设置栏位高度
        if j == 4:
            tall_style = xlwt.easyxf('font:height 800;') #设置字体高度
            row0 = table.row(j)
            row0.set_style(tall_style)
        elif j == 1:
            tall_style = xlwt.easyxf('font:height 500;') #设置字体高度
            row0 = table.row(j)
            row0.set_style(tall_style)
        else:
            tall_style = xlwt.easyxf('font:height 300;') #设置字体高度
            row0 = table.row(j)
            row0.set_style(tall_style)
        
    a = i[0]
    #print(a)
    cur.execute("select * from %s" % a)
    col_name_list = [name[0] for name in cur.description]
    #print(col_name_list)
    
    cur.execute('show create table `%s`;' % a)
    b = cur.fetchall()
    #print(b)
    #print(len(b))
    for k in b:
        #excel内的列数
        excel_col = 1
        table_name = k[0]
        print(table_name)
        structure = eval(repr(k[1]).replace('\\n  ', ''))
        print(structure)
        pattern = re.compile(r"[(]`.*\n[)]")
        result = pattern.findall(structure)
        print(result[0])
        #result[0]前后的括号去掉 后面额外去掉
        result[0] = result[0][1:-2]
        result[0] = result[0].replace('\n','')
        
        print("删去了括号和换行",result[0])
        #字符串分割，但是纯用逗号分割会引起问题，先使用正则表达式更改一些不应该被分开的逗号 删除primary key 所在行
        
        #替换掉所有括号里的逗号
        split_pattern = re.compile(r' FOREIGN KEY [(]`[^()]*?,[^()]*?[)]')
        foreign_key_things = re.findall(split_pattern, result[0])
        
        print("foreign_key_thing", foreign_key_things)
        #先找到decimal(,)结尾括号的位置
        for foreign_key_thing in foreign_key_things:
            print("替换逗号") 
            replaced_foreign_key_thing = foreign_key_thing.replace(',','|')
            result[0] = result[0].replace(foreign_key_thing, replaced_foreign_key_thing)
            
        print("替换foreign_key逗号后", result[0])
        
        split_pattern = re.compile(r' REFERENCES [(]`[^()]*?,[^()]*?[)]')
        references_things = re.findall(split_pattern, result[0])
        
        print("references_thing", references_things)
        #先找到decimal(,)结尾括号的位置
        for references_thing in references_things:
            print("替换逗号") 
            replaced_references_thing = references_thing.replace(',','|')
            result[0] = result[0].replace(references_thing, replaced_references_thing)
            
        print("替换references逗号后", result[0])
        
        split_pattern = re.compile(r' decimal[(]\d*,\d*[)]')
        decimal_things = re.findall(split_pattern, result[0])
        
        print("decimal_thing", decimal_things)
        #先找到decimal(,)结尾括号的位置
        for decimal_thing in decimal_things:
            print("替换逗号") 
            replaced_decimal_thing = decimal_thing.replace(',','|')
            result[0] = result[0].replace(decimal_thing, replaced_decimal_thing)
            
        print("替换decimal逗号后", result[0])
        
        ## float 精度的逗号问题
        split_pattern = re.compile(r' float[(]\d*,\d*[)]')
        float_things = re.findall(split_pattern, result[0])
        
        print("float_thing", float_things)
        #先找到decimal(,)结尾括号的位置
        for float_thing in float_things:
            print("替换float逗号") 
            replaced_float_thing = float_thing.replace(',','|')
            result[0] = result[0].replace(float_thing, replaced_float_thing)
            
        print("替换float逗号后", result[0])
        
        #最短匹配 PRIMARY(,)
        split_pattern = re.compile(r'PRIMARY KEY [(]`.*?`,`.*?`[)]')
        primary_things = re.findall(split_pattern, result[0])        
        print("primary_thing", primary_things)
        for primary_thing in primary_things:
            print("删除primary语句")
            result[0] = result[0].replace(primary_thing, '')
        
        print("删除primary语句后", result[0])
        
        #最短匹配 KEY(,)
        split_pattern = re.compile(r'KEY `.*?` [(]`[^()]*?`,`[^()]*?`[)]')
        key_things = re.findall(split_pattern, result[0])        
        print("key_thing", key_things)
        for key_thing in key_things:
            print("删除key语句")
            result[0] = result[0].replace(key_thing, '')
        
        print("删除key语句后", result[0])
        
        columns_information = result[0].split(',')
        
        
        #记录表项名字和所在excel列数的字典
        column_name_excel_col_dict = {}
        print('columns_informations', columns_information)
        #remark_flag = False
        
        #记录自增长的列的字典
        self_increasement_name_excel_col_dict = {}
        
        for column_information in columns_information:
            if len(column_information) == 0:
                continue
            #是字段
            if column_information[0] == '`' or column_information[0] == '(':
                print('columns_information', column_information)
                #"|"替换为逗号
                column_information = column_information.replace('|',',')
                
                
                #拿到列名
                column_name_pattern = re.compile(r'`.*` ')
                column_name = column_name_pattern.findall(column_information)[0]
                column_name = column_name[1:-2]
                table.write(0, excel_col, column_name, setStyle('等线', height=220, bold=False))
                print("column_name", column_name)
                
                #对一些常见的进行翻译
                if column_name in chinese_column_name:
                    table.write(1, excel_col, chinese_column_name[column_name], setStyle('等线', height=220, bold=False))
                else:
                    table.write(1, excel_col, '', setStyle('等线', height=220, bold=False))
                    
                #写入字典
                column_name_excel_col_dict[column_name] = excel_col
                #字母(数字)
                column_type_pattern1 = re.compile(r' [a-zA-Z]*[(].*[)]')
                column_type = column_type_pattern1.findall(column_information)
                print(column_type)
                #匹配到了 证明是有长度的列
                if len(column_type) != 0:
                    column_type = column_type[0][1:]
                    print("column_type", column_type)
                    table.write(2, excel_col, column_type, setStyle('等线', height=220, bold=False))
                #没匹配到
                else:
                    # `+空格+字母+空格|字符串尾
                    column_type_pattern2 = re.compile(r'` [a-zA-Z]*(?:$| )')
                    column_type = column_type_pattern2.findall(column_information)[0]
                    column_type = column_type[2:]
                    print("column_type", column_type)
                    table.write(2, excel_col, column_type, setStyle('等线', height=220, bold=False))
                
                #检查是否包含DEFAULT NULL 或者 NOT NULL 或者 AUTO_INCREMENT
                if 'NOT NULL' in column_information:
                    #非空
                    table.write(3, excel_col, '是', setStyle('等线', height=220, bold=False))
                else:
                    table.write(3, excel_col, '否', setStyle('等线', height=220, bold=False))
                    
                if 'AUTO_INCREMENT' in column_information:
                    #自增长
                    print(column_name,"自增长")
                    table.write(4, excel_col, '自增长', setStyle('等线', height=220, bold=False))
                else:
                    table.write(4, excel_col, '', setStyle('等线', height=220, bold=False))
                #下一列
                excel_col += 1
            #这句表示外键约束        
            elif column_information[:10] == 'CONSTRAINT':
                #替换回去逗号
                column_information = column_information.replace('|',',')
                print(column_information)
                column_remark_pattern = re.compile(r'REFERENCES `.*` ')
                column_remark = column_remark_pattern.findall(column_information)[0]
                column_remark = column_remark[12:-2]
                print("column_remark", column_remark)
                foreign_key_name_pattern = re.compile(r'FOREIGN KEY [(]`.*`[)] REF') 
                foreign_key_name = foreign_key_name_pattern.findall(column_information)[0]
                print("foreign_key_name", foreign_key_name) 
                foreign_key_name = foreign_key_name[14:-6]
                print("foreign_key_name", foreign_key_name)
                #可能是多个foreign key
                if foreign_key_name in column_name_excel_col_dict:
                    table.write(4, column_name_excel_col_dict[foreign_key_name], column_remark + '外键', setStyle('等线', height=220, bold=False))
                else:
                    #逗号分割
                    foreign_key_names = foreign_key_name.split(',')
                    print('foreign_key_names',foreign_key_names)
                    for nm in foreign_key_names:
                        nm = nm.replace('`', '')
                        nm = nm.replace(' ', '')
                        print('nm', nm)
                        table.write(4, column_name_excel_col_dict[nm], column_remark + '外键', setStyle('等线', height=220, bold=False))
            
        for y in range(excel_col):
            table.col(y).width = 256 * (30 + 1)

file.save(r'test.xls')
#关闭连接
conn.close()



