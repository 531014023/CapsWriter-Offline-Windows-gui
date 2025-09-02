![](https://aliyun.93dd.top/picgo/20250901171510426.png)

# 软件介绍
CapsWriter-Offline是CapsWriter 的离线版，一个好用的 PC 端的语音输入工具。但是离线版是控制台运行，不方便使用，我就用deepseek写了个python的gui图形界面工具来管理。
# 下载地址
[HaujetZhao/CapsWriter-Offline: CapsWriter 的离线版，一个好用的 PC 端的语音输入工具](https://github.com/HaujetZhao/CapsWriter-Offline) ,使用方法主要看这个，介绍的最全面。

[夸克下载链接](https://pan.quark.cn/s/caa29c83e985)，这是我打包exe后的下载地址，里面有打包工具的代码和打包后的文件以及gui图形界面的使用介绍。软件纯绿色，不需要安装，解压放置模型即可使用。

# 使用方法
只需要下载`CapsWriter-Offline-Windows-gui-exe.win-amd64-3.11.zip`和`models.zip`这两个压缩包，解压到自己新建的文件夹，models.zip解压放置到`exe.win-amd64-3.11\CapsWriter-Offline-Windows-64bit`下面，运行`exe.win-amd64-3.11`下的`.exe`文件即可使用。

# 自行打包使用方法
```
python setup.py build
```
此命令会在当前目录生成build文件夹，其中就是一个完整的应用，执行exe即可打开使用，需要注意的是在使用命令前，需要先将[github](https://github.com/HaujetZhao/CapsWriter-Offline)程序`CapsWriter-Offline-Windows-64bit`放到当前目录供打包使用。打包后执行的是`CapsWriter-Offline-Windows-64bit`下的start_all.vbs，此文件是自定义的，就是隐藏黑窗口命令：
```
CreateObject("Wscript.Shell").Run "start_server.exe",0,False
CreateObject("Wscript.Shell").Run "start_client.exe",0,False
```
如果不是用的我提供的程序就需要自己添上start_all.vbs文件。

模型也需要去[github](https://github.com/HaujetZhao/CapsWriter-Offline)的发布页面下载model.zip，下载后放置到打包后的exe文件同级的`CapsWriter-Offline-Windows-64bit`目录下即可使用。
## 打包前目录结构
![](https://aliyun.93dd.top/picgo/20250901163957774.png)

## build目录结构
build下面只有一个文件夹`exe.win-amd64-3.11`，下面的结构如下：
![](https://aliyun.93dd.top/picgo/20250901164041139.png)

## 模型文件放置位置
![](https://aliyun.93dd.top/picgo/20250901164222308.png)
