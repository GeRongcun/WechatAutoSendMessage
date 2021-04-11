**摘要：**利用Pywinauto库，模拟键盘鼠标操作微信，自动发送提醒消息，提醒大家完成“青年大学习”。

## 引言

### 问题的提出

进入研二后，因为科研任务和就业压力，大家参加“青年大学习”的积极性不是很高。部分同学需要多次提醒，才会完成“青年大学习”。
提醒的方式有很多种：1)在群里提醒，并@全体成员，缺点是影响到已学习过的同学；2）在群里提醒，不@全体成员，缺点是有些同学会设置免打扰，提醒效果不好；3）在群里提醒，并@没有学习的同学，缺点是需要输入同学姓名，比较费时间；4）私聊提醒，缺点同样是费时间。

为了兼顾提醒效果和时间消耗量，我利用Pywinauto库，模拟键盘鼠标操作微信，自动发送提醒消息，提醒大家完成“青年大学习”。操作微信，既可以在群里提醒，并@没有学习的同学，又可以私聊提醒。

### 实现思路

一、获得学习情况信息

咱们班记录学习情况的工具是腾讯文档，学习完成后，在腾讯文档对应列填写“已学习”。

![](https://www.gerongcun.xyz/blog/2021/fa6383fb/腾讯文档2.png)

1）使用Selenium库模拟鼠标键盘操作，登录腾讯文档，将表格导出成Excel文档。

2）使用Pandas、Numpy库处理Excel文档，遍历全部同学，如果该同学的“个人申报”列为空，则该同学没有学习，进入获得没有学习同学的名单。

二、操作微信，进入聊天界面

1）利用Pywinauto库，连接到微信，定位到微信左上角的搜索框。

![](https://www.gerongcun.xyz/blog/2021/fa6383fb/微信搜索框.png)

2）在搜索框里输入班级群名或者同学名字，进入聊天界面。

![](https://www.gerongcun.xyz/blog/2021/fa6383fb/搜索框输入群名2.png)

三、生成提醒消息，发送消息

1）利用Template库，设计模板，自动生成提醒消息，发送消息

![](https://www.gerongcun.xyz/blog/2021/fa6383fb/生成提醒消息.png)

2）整合以上功能，利用Time库，在一天中的几个时间点自动发送提醒消息。

### 具体效果

群发消息：  
![](https://www.gerongcun.xyz/blog/2021/fa6383fb/效果图1.png)

私发消息：
![](https://www.gerongcun.xyz/blog/2021/fa6383fb/效果图2.jpg)

### Pywinauto学习资料

Pywinauto是一组用于自动化Microsoft Windows GUI的python模块。 最简单的是，它允许您将鼠标和键盘操作发送到窗口对话框和控件。Selenium和Pywinauto都是自动化工具，Selenium适用于Web浏览器，而Pywinauto适用于PC端应用程序。

我整理了一些Pywinauto学习资料：  
[Pywinauto中文文档](https://www.kancloud.cn/gnefnuy/pywinauto_doc/1193035)  
[python基于pywinauto实现PC端自动化 python操作微信自动化](https://www.cnblogs.com/xp1315458571/p/13892205.html)  
[Pywinauto笔记ByRowingMan_v1.0.docx](Pywinauto笔记ByRowingMan_v1.0.docx)
