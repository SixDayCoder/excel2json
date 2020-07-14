# 使用方法

使用方法 : python excel2json.py input_excel_path output_json_path
input_excel_path : 输入excel文件
output_json_path : 输出的json文件

# Excel表的格式

Excel的表格式必须满足excel_format.png图示的格式

第一行 : 描述字段名称,其中第一列必须为Id
第二行 : 描述每个字段对应的类型,目前仅支持：
		 (1)INT
		 (2)INT64
		 (3)STRING
		 (4)FLOAT
		 (5)DATE(yyyymmddddhhmmss)
第三行 : 每个字段的名称注释


# 关于数组

excel2json支持在表格中按照一定规则填写后到出json的数组，例如:

AwardId_1	AwardCnt_1	AwardId_2	Award_Cnt_2
INT		  	INT		 	INT       	INT
物品Id1	  	物品数量1	物品Id2		物品数量2
1			2			3			4

生成的json的格式为：

"AwardId":[1, 2],
"AwardCnt" : [2, 4]

即以_为标记，标记该字段是数组字段

