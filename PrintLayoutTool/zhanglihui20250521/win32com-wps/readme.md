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

## 若 PowerShell 禁止运行脚本

Windows 默认会限制 `.ps1` 脚本执行，可能出现：

```text
无法加载文件 ...uv-run.ps1，因为在此系统上禁止运行脚本。
```

**解决办法（二选一）：**

**方案 A：仅本次运行绕过（推荐，无需管理员权限）**
```powershell
powershell -ExecutionPolicy Bypass -File .\uv-run.ps1
```

**方案 B：为当前用户永久允许本地签名脚本**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

查看当前策略：
```powershell
Get-ExecutionPolicy
```

设置完成后，即可直接右键“使用 PowerShell 运行” `.ps1` 脚本，或执行 `.\uv-run.ps1`。