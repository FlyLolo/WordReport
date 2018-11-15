# WordReport
   经常会有生成word版本的报告的例子, 例如体检报告;
   这是一个通用的Word报告生成程序, 将Dataset中的数据按照设定的规则写入模板,支持导出docx、doc和PDF.
   Dataset指提供数据, 报告中的样式(例如字体、表格宽度、对齐方式等)全部在word模板中可视化定制.
   
   功能已经实现, 一些细节暂未处理, 比如对于模板配置错误时的友好提示和日志.
   速度有点慢, 据说OpenXML会好很多,但据说只支持word的docx版本,导出doc和pdf不知道是否可行, 有时间深入研究一下.
   
   
# Demo
做了个将Sql Server数据库的表结构导出的demo.

# 使用方法
1.将文档需要的数据整理到一个Dataset(本文以sql server做的例子, 您可以使用其他的).
2.新建一个docx格式的word文档, 通过新建书签的方式划定区域,规则见下文.
3.根据书签中的描述, 系统自动将数据写入相应位置.

# 模板规则
  见我的博客https://www.cnblogs.com/FlyLolo/p/WordReport.html. 当时写的有点粗糙, 有时间整理一下
 
