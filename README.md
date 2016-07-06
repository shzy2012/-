# -Web services 半自动化测试程序

目录结构：

   Tester.cs --主程序
   
   WebServices.xml --接口配置文件
   
   packages.config --项目运行所需要引用的库


程序测试设计思路： 通过在 WebServices.xml文件中配置web service的Url和参数名称和类型，以及参数的默认值， 通过C#解析xml文件，发起Http请求，最后生成测试结果.

过程图：

 配置 xml--> C# --> 读取xml --> 生成http参数 -->发起 HttpWebRequest -->返回结果 HttpWebResponse -->分析 -->报告
