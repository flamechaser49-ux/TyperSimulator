# 文件打字模拟器

把内容像真人打字一样**逐字敲进"光标点击的输入框"**——网页输入框、聊天窗口、记事本、VSCode、Word 都可以。

## 运行

- **双击 `打字模拟器.exe`**（推荐）：免 Python，任何 64 位 Windows 10/11 都能跑。
  运行时会在 exe 旁边生成 `_MEIxxxx` 临时文件夹（不碰 C 盘临时目录），退出时自动删除。
- 有 Python 时也可以：`python TyperSimulator.py`
- 源码自测：`python -X utf8 TyperSimulator.py --selftest`

## 拷给其他电脑

只需拷 `打字模拟器.exe` **这一个文件**。注意：

- 首次运行可能弹 SmartScreen"已保护你的电脑"→ 点"更多信息"→"仍要运行"（新文件没有信誉记录的正常现象）
- 数字签名为自签名证书（签名者 Ruitariel）：文件属性里能看到签名者，但因为不是公信 CA 签发，会显示"不受信任"——不影响运行

## 支持的文件与编码

| 类型 | 后缀 | 编码 |
|---|---|---|
| 文本 | `.txt` | UTF-8 / UTF-8 BOM / UTF-16 / GBK（自动识别） |
| C 语言 | `.c` `.h` | 同上 |
| C++ | `.cpp` `.cc` `.cxx` `.hpp` | 同上 |
| Python | `.py` | 同上 |
| Word | `.docx` | 自动提取正文 |

ASCII 是 UTF-8 的子集，天然支持。旧版 `.doc` 不支持，程序会提示"另存为 .docx 或 .txt"。

## 使用方法

**方式一：选文件** —— 点 **选择文件…**，内容进入左侧内容框（框里显示识别出的编码）。

**方式二：直接粘贴** —— 复制任意文本，点内容框下的 **粘贴剪贴板**（或直接在框里 Ctrl+V / 手动编辑），框里是什么就打什么。

然后：

1. 调好参数：**每秒字符数**（默认 15，上限 500）、**模拟抖动**（更像真人）、或勾 **硬输入**（不限速全速直灌，勾上后速度和抖动失效）
2. 点 **开始 (F6)**（或直接按 F6）
3. 倒计时期间，**用鼠标点一下你要打字的那个输入框**（让光标停在那里）
4. 倒计时结束，内容就开始逐字敲进去

想先在本程序里看看效果：把模式切到 **内置演示框**，内容会打在右侧演示框里。

## 快捷键（全局，任何窗口里都有效）

| 键 | 作用 |
|---|---|
| F6 | 开始 / 暂停 / 继续（倒计时中按 = 取消） |
| Esc | 停止 |

## 文件结构

```
typing_simulator/
├── 打字模拟器.exe     独立可执行（分发只需这一个文件）
├── TyperSimulator.py  源码（单文件，改功能后可重新打包）
└── README.md
```

## 改完源码重新打包

```bash
python -m PyInstaller --onefile --windowed --runtime-tmpdir . --name TyperSimulator TyperSimulator.py
mv TyperSimulator.exe 打字模拟器.exe
# 如需重新签名（证书在当前用户证书库，CN=Ruitariel）：
powershell -Command "$c = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object {$_.Subject -eq 'CN=Ruitariel'}; Set-AuthenticodeSignature .\打字模拟器.exe -Certificate $c"
```

## 注意事项

- 按键以 Unicode 方式发送，中文、全角标点、emoji 都支持；文件里的换行会按回车键。
- 打字时自动把系统定时器精度提到 1ms，高速度（几百字/秒）才能真正跑满。
- 速度极快或开硬输入时，个别接收慢的程序可能丢字，遇到就把速度调低一些。
- **管理员权限运行的窗口**（如某些系统程序）收不到普通程序发的按键，需要本程序也以管理员身份运行。
- 部分游戏使用 DirectInput 读键盘，可能收不到模拟按键。
- F6 在个别程序里本身有快捷键（如浏览器切换焦点），如冲突可改源码里的 `VK_F6` 常量后重新打包。
