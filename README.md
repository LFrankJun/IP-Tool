# IP-Tool
IP工具代码

	V3.2 : 添加对Word文档中只有“权利要求书”或者“说明书”或者“说明书摘要”的质量检查（不带任何页眉或者关键字）；

	V3.3： 
	1）删除多余代码 ；

	2）修改缺乏引用基础的代码，当进行缺乏引用基础的判断时，忽略某些标点符号（&，“，”，‘，’，《，》，{，}，]，（，），[）;
	
	v3.4： 仅检查权利要求时，检查结果页面没有进行最大化展示，对此bug进行修改；
	v3.5:  修复标题中带有空格或者其他标点符号，无法获取到标题的异常情况；
	v3.6:  添加附图标记说明是表格的情况；
	v3.7:  添加附图标记PDF的cid字符的处理以及“附图标记说明”关键字和附图标记的相关字符串处于同一个段落情况下的处理;
	v3.8:  1）适配附图标记中存在表格，同时文档其他区域也存在表格的情况； 2）优化获取附图标记的名词和标号；
	v3.9:  1)修复判断“缺乏引用基础”时，因为getforma函数导致的出现-1的情况； 2）添加startIndex和endIndex的初始化
