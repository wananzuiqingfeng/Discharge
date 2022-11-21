由于工作需要统计 BatteryMon 生成的电池曲线数据文件到 Excel 中，

项目需要长时间处理大量重复性的工作，特此开发此工具节省时间。

这是一个利用 Python + openpyxl 开发的一款自动化脚本工具，

帮助你将电池曲线数据快速录入至 Excel 文件。

## 安装

```
Tips: 已安装 Python3.8 及以上版本 + openpyxl 框架可跳过安装步骤!
```

Discharge 目录文件如下:

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/1.png)

双击打开 EnvironmentInstall 文件夹，如下:

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/2.png)

运行 python-3.9.6-amd64.exe 安装 Python3.9(版本在 Python3.8及以上即可)

勾选 Add Python 3.9 to PATH 后再点击 Install Now

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/3.png)

安装 Python 后运行 Pip.bat 安装 openpyxl 第三方库:

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/4.png)

结果如下(不出现红色字体的错误提示即为安装成功):

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/5.png)



## 使用

右击打开 Discharge 目录下的 Run.bat 文件进行编辑:

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/6.png)

Run.bat 内容如下:

```
@C:\Users\SuperPig\AppData\Local\Programs\Python\Python310\python.exe discharge.py %*
@pause
```

```
Tips: 狗作者在编辑教程的时候电脑的 Python 版本是 3.10，根据此电脑的 Python 安装路径替换即可
```

Win + R 启动 cmd 窗口，输入指令：

```
py -op
```

获取此电脑 Python 的安装路径，将此电脑的 Python 安装路径替换掉 Run.bat 中的默认路径

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/7.png)

因公司对电池曲线标准为两台设备做四个循环，舍弃第一轮数据。

BatteryMon 工具生成的文件默认命名规范为前缀 f 表示放电曲线，c 表示充电曲线，后缀用数字表示第几轮

```
Tips: 因此使用该工具时请保证数据文本文件的名称符合规范，否则运行会报错!!!
```

将数据文本文件放入到 Discharge 目录下，如下所示:

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/8.png)

运行 Run.bat 即可收录当前设备的数据到 Excel 表格，成功运行结果如下:

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/9.png)

```
Tips: 充放电数据的不完整可能使程序无法正常执行，请注意输出信息来追寻错误，或者联系狗作者寻求帮助。
```

生成的数据存放在 Result 文件夹下，每次操作会根据当前日期生成结果文件夹:

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/10.png)

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/11.png)

文件夹 1 表示当前操作设备为 #1，存放的文件就是当前操作的数据文件以用于备份

Excel 文件则是最终需要的报表数据，对于收录 #1 设备的电池曲线数据后，该文件如下:

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/12.png)

公司标准要求做两台设备的数据，因此再将 #2 的数据放入 Discharge 目录运行 Run.bat 即可:

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/13.png)

```
Tips: 作为演示，一台设备的数据被我们重复使用，因此设备 #1、2 的数据是一样的，不要在意 ^-^
```
### 演示

![](https://github.com/wananzuiqingfeng/Discharge/blob/master/Images/Demo.gif)

## 联系狗作者

```
企鹅: 1424286760
```
