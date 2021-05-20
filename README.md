# CdrExtension-ABH

    [CorelDraw GMS Extension And WPF Installer]
    【CDR GMS 扩展插件 以及 安装包】

## 说明 / Information

- Full Name: Albertiy's Birthday Hat Extension
- Language: VBA
- Installer Language: C# WPF

## 注意 / Notice

.GMS file 密码/Password：zha123

To use the GMS source code, `Alt+F11` open the CDR macro editor, and then import them into a global macro project. (You can open the 'Draw/GMS' directory exists in where you installed the CorelDraw Software, then copy a .gms file as your new gms project.

如果想要使用GMS源码，首先打开安装的 CorelDraw 软件，`Alt+F11` 打开宏编辑器，在 GMS 项目上导入。可以到 CDR 安装目录的 Draw/GMS 文件夹中，复制一个 .gms 文件，作为你自己的新工程。 

- CDR X7 GMS 文件夹位置：%安装目录%\Corel\CorelDRAW Graphics Suite X7\Draw\GMS
- CDR X4 GMS 文件夹位置：%安装目录%\Corel\CorelDRAW X4\Draw\GMS

X4 and X7 GMS is not in common use, and some method is different. But you can import from source code and modify it.

X4 和 X7 的 GMS 格式不是通用的，甚至可能导致崩溃。并且，一些方法也不同。但是，可以通过导入源码并修改来解决。比如 X7 中的 createBoundary()，X4 中需要使用替代的方法。

## 2021-05-18 16:52:08 修改：

按照不同的版本所需插件不同，分为 X4 版本 和 X7 两个版本，路径中带有 X4，则复制 X4 版本的插件，否则都是用 X7 版本插件。

要复制的文件在 X7/GMS 和 X4/GMS 文件夹中，打包成 exe 前需按照版本复制到 ABHInstaller/InstallResources/X4 和 X7 文件夹下（没有自己建），然后修改这些文件的生成操作为 Resource。

## 2021-05-20 17:22:47 修改：

为 ABHInstaller 添加了 .zip压缩文件 作为资源并自动解压到安装位置的功能（Thanks SharpZipLib）。节省了每个文件都要写入资源列表的烦恼。为此添加了几个工具类。