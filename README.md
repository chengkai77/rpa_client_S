##### 项目环境依赖

    1. 引用文件环境(pyinstallerenv_client_ocr)

##### 项目目录

```
 ├── rpa_client                     //项目文件夹
 │
 ├── icons                          //图标文件
 │
 ├── Lbt                            // 功能模块代码文件
 │
 ├── local                          // 项目国际化语言配置文件
 │  
 ├── log                            // 日志文件
 │                  
 ├── plugin                         // 编译文件
 │                  
 ├── utils                          // 工具文件
 │                 
 ├── xxx_target.py                  // 功能模块文件  
 │
 ├── tray.py                        // 机器人推盘文件 
 │
 ├── ChocLead.spec                  // pyinstall 打包编译文件
 │
 └── ChocLead.py                    // 项目入口文件

```

##### Log 分模块存储对应编号

```angular2html
1. 客户端主程序    Robot Client    10000
2. Access模块    Accsee    10001
3. AI模块    AI    10002
4. Excel模块    Excel    10003
5. 文件模块    File    10004
6. 通用模块    General    10005
7. 人机交互模块    HcI    10006
8. 邮件模块    Mail    10007
9. 主功能    MainFunction    10008
10. 鼠标与键盘模块    MouseKeyborad    10009
11. SAP模块    SAP    10010
12. 网页模块    WEB    10011
13. 自定义编程    SelfDevelopment    10012

... 没有定义默认使用10000
```

* 不同地方使用 日志logger操作填入对应编号，模块为10000
