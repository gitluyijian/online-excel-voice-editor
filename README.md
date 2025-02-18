# online-excel-voice-editor

#### 介绍
在线excel语音输入编辑器，已支持ipad浏览器，帮助用户在进行试验或者工作工程中无法腾出双手进行excel表单录入数据时，通过语音对话能直接完成excel数据录入，最新执行导出

#### 软件架构
软件架构说明


#### 安装教程

1.  python版本3.11
2.  打开到项目目录，命令行打开执行 pip install -r requirements.txt
    ![输入图片说明](https://foruda.gitee.com/images/1739840186084628713/e9c02bd2_2287099.png "屏幕截图")
3.  命令行执行 python app.py
    ![输入图片说明](https://foruda.gitee.com/images/1739840218726526070/6bbd167c_2287099.png "屏幕截图")

#### 使用说明

1.  浏览器访问进行注册登录
    ![输入图片说明](https://foruda.gitee.com/images/1739840320485462358/8c4cc2cb_2287099.png "屏幕截图")
2.  登录后选择自定义excel模板文件，点击上传
    ![输入图片说明](https://foruda.gitee.com/images/1739840393643031346/711463ee_2287099.png "屏幕截图")
3.  开启语音输入
    ![输入图片说明](https://foruda.gitee.com/images/1739840427356054549/ffe8d960_2287099.png "屏幕截图")
4.  通过语音输入自动填入数据（目前使用的是系统自带的webkitSpeechRecognition，对专业性的语音识别准确率比较低）  
    ![输入图片说明](https://foruda.gitee.com/images/1739840593001958709/165a199c_2287099.png "屏幕截图")
5.  系统预设单元格移动识别语音文字
    ![输入图片说明](https://foruda.gitee.com/images/1739840680126216222/aabba085_2287099.png "屏幕截图")
