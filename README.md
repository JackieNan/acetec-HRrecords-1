# acetec-examrecords-1

## 功能描述

1、针对不同种类的清单（白班/产线）中出现的异常时间（迟到、打卡异常、在公司外时间过长等），根据原始“请假“、“异常”记录，初步筛选出无请假、异常记录的个体，供人事部门人员进行审阅；

## 使用说明

1、在使用程序之前，需要将两个清单进行简单处理：在`查找和替换`中，将填充颜色为“FF0000”和“FFA500”的单元格替换为相同的填充颜色（刷新）；

2、将原始记录表格另存为xlsx格式的表格；

3、打开程序，按照“清单——记录”的顺序拖入弹出框内，一次只拖入一种清单，点击”运行“；

4、等待，在程序相同目录下出现一个叫做“结果”的excel文件，即为处理完毕的表格；

注意：在处理完一个表格后，**一定要将“结果”文件改名**，否则下一生成表格会覆盖上一次的结果！

​			在程序运行期间，**不要打开**“结果”文件！

​			程序同目录下的"data1"和"data2"文件为程序生成结果时处理数据用到的文件，在程序运行结束后**可以**删除。

