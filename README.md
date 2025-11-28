# 皎皎助手 - 二重螺旋后台挂机脚本

⚠️ 本脚本基于 pywin32 调用 Windows API 执行截图与输入操作，无内存注入等高危操作！⚠️ 

2025.11.28 27号官方大面积封禁每日在线时长较长的脚本用户，同时经群聊统计，win32脚本被封禁比例较大，因此本项目更新停止上传。

~~期末周没有时间写，密函选择还有部分识图的实现有bug，假期再搞吧~~

![脚本窗口示例图片](Guiexample.png)

# 环境准备
```powershell
# 安装所需依赖库
pip install -r requirements.txt
运行脚本

# 启动主程序
python main.py

# 使用pyinstaller打包脚本主程序
pip install pyinstaller
python -m PyInstaller "ur dir\JiaoJiao\mainGuiRe.py" -F -w --name "name" --collect-submodules logic --collect-data ttkbootstrap --distpath "ur dir\JiaoJiao\dist" --workpath "ur dir\JiaoJiao\build" --specpath "ur dir\JiaoJiao"

# 使用pyinstaller打包操作录制器
python -m PyInstaller "ur dir\JiaoJiao\recorder.py" -F -w
```

```markdown
## 项目目录结构
- jiaojiao/  # 根目录
  - control/  # 按键模板图
    - json/  # 移动序列文件目录
      - 55mod/
      - juesemihan/
      - wuqimihan/
  - logic/  # 各模式循环逻辑
    - 55mod.py
    - juesemihan.py
    - wuqimihan.py
  - map/  # 地图资源目录
    - 55mod/
    - juesemihan/
    - wuqimihan/
  - config.json  # 用户设置保存文件
  - jsontest.py  # JSON操作序列测试用
  - main.py  # 主程序入口
  - recorder.py  # 操作录制器
  - test.py  # 非焦点窗口截图测试脚本
  - test2.py  # 非焦点窗口输入操作测试脚本
  - .....
```

# 功能说明

当前已实现功能：
55 夜航模式挂机
驱离武器密函模式挂机

# 可能遇到的问题
Q:地图，按键等识别失败。
 A:由于每个设备的分辨率等不同，若脚本一直无法识别按键和地图，你可能需要更换./map和./control的特征图，请检查游戏内设置：设定16：9，1920x1080，清晰度中或高。PC设备设置：缩放100%

Q：脚本运行后无反应。
 A：请使用管理员模式开启

Q：不是以上问题，就是运行不了。
 A：请描述问题，发送日志并标明来意至：1484413790@qq.com

# ToDo
扩展更多游戏内模式支持.

~~优化图像识别逻辑，提升成功率.~~

增强失败补救机制（如重新点击、退出重开等).

增加自定义快捷键功能.支持组合键操作（如一键螺旋飞跃二段跳冲刺快速跑图，一键复位等）.

~~优化Gui界面，提升人性化体验.~~

增加各种抖动之类的安全防检测措施




# 更新日志

2025.11.7
使用ttkbootstrap优化了界面gui，以圆角卡片形式分布功能

优化了日志文本，增加emoji更好理解

优化了recorder，可以手动选择地图模板图和json保存位置

优化了地图识别逻辑，删除所有阈值限制，三特征匹配并把三特征置信度相加选最高者，这样调整之后地图识别成功率很高且不怎么误判，如果长期测试后还是会出现误判问题，考虑增加roi蒙版🤔

解决了部分按键的无法后台点击问题，通过深度子控键遍历，添加了CHILItest.py用于测试用例和实现代码留存。


