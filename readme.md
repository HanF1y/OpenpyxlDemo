一、获得表格，创立标志 字符串型队名tname1、tname2、整数型分数主队场分goal1、客队场分goal2、主队女将比分mfegoal、客队女将分fegoal、局分wgoal、负场数fgoal、布尔型主将胜负master、

二、按行查找“主队胜/负”，胜场记录局分wgoal，负场分数记做fgoal，比较两者关系（4:0和3:1都是胜方3分，负方1分，2:2的话，以第一台胜方为最终胜方，积2分，败方积1分）2：2时查看“第一台（主将）”所在行胜负

三、按列找“性别”，若为“女”，按行查找胜负，记主队女将比分mfegoal、客队女将分fegoal

四、查找队名

五、按列查找队名tname1、tname2

六、将对应数据写入积分表格内

七、使用说明：

1、test1为比赛结果表格，每次只读入一个表格，需修改路径

2、test2为积分表，所有积分初始化为0，需修改路径，此后每次存储时需要在最后一行处更改文件名

3、自动查找表格内容，与表格所在位置无关，运行过程约需1s

4、只用于交流学习
