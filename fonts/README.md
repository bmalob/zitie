# fonts/ 字体目录

把你自己的字体文件（`.ttf` / `.otf` / `.ttc`）放进这个目录，然后运行：

```bash
python3 zitie.py --list-fonts
```

即可看到并选用这些字体，例如：

```bash
python3 zitie.py content.txt --font "田英章硬笔楷书简体" --pdf
```

也可以不放在这里，直接用 `--font-file /路径/某字体.ttf` 指定任意位置的字体。

## 随附字体

- `LXGWWenKai-Regular.ttf` —— **霞鹜文楷**（LXGW WenKai），SIL Open Font License 1.1，
  免费可商用，许可证见同目录 `OFL-LXGW-WenKai.txt`。用 `--font wenkai` 选择。

## 版权说明

田英章、庞中华、司马彦、吴玉生等书家字体**多为商业字体**，版权属于字体厂商或作者，
因此**不随本项目分发**。个人练字一般可自行搜索下载后放入本目录使用；
如需商用、分发或公开发布，请购买正版授权。

字体相关的版权问题由使用者自行承担，与本项目代码无关。
