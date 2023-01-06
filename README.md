# ht-Pandas-mess
######
## This one is just for the record
######
io: scorecard文件路径
VS_input：名字-ValueStream文件路径 需要在excel里添加需要查找的主管的名字和Value-Stream
PID_input_path：主管级PID审核分值统计表的文件夹路径
ldl_input_path：主管领导力特质评分表的文件夹路径

table_num：scorcard中一个主管子表的行数+2 
month_num：需要计算的月份，范围1~12


x_min,x_max=30,60
y_min,y_max=0,40
x_d=(x_max-x_min)/3
y_d=(y_max-y_min)/3
与九宫格九个格子划分有关，xmin,xmax为九宫格x轴的最小值和最大值，ymin,ymax为九宫格y轴的最小值和最大值，

如果要更改指标计算逻辑，找到对应的指标计算函数在里面修改即可
