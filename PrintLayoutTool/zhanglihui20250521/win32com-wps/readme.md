安装步骤：

1.安装uv环境
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

<https://uv.doczh.com/getting-started/installation/#__tabbed_1_2>

2.进入py脚本目录，创建虚拟环境，并激活
```powershell
cd PrintLayoutTool\zhanglihui20250521\win32com-wps
uv venv
.venv\Scripts\activate
```

3.安装依赖包（使用清华镜像加速）
```powershell
uv pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

4.运行py脚本

```powershell
python .\XFormz_zlh_v7.py
```

5.通过ps1脚本运行py脚本

```commandline
powershell -ExecutionPolicy ByPass -c "uv run .\XFormz_zlh_v7.py"
```

把这条命令写入ps1脚本中，然后可以在窗口右键以“使用PowerShell运行”的方式鼠标操作运行ps1脚本