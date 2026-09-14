# PPT Animation Coach

这是一个专注于 PPT 图片轮播动画的 Codex Skill。

它可以接收用户提供的 5 张图片，在安装了 Microsoft PowerPoint 的 Windows 电脑上生成可播放、可编辑的图片轮播演示稿：

- 五张图片组成圆角卡片轮播布局
- 中间为大图，两侧为中小图
- 图片沿顺时针方向轮换
- 使用 PowerPoint「平滑」切换实现连续动画
- 保留图片框的对象对应关系，方便后续编辑
- 自动在演示稿备注中写入制作原理和操作要点

## 使用场景

当用户希望把 5 张图片制作成类似图片画廊、作品展示、旅行相册或产品轮播效果时，可以使用这个 Skill。

## 文件说明

- [`SKILL.md`](SKILL.md)：Skill 的主要使用说明
- [`references/technique-index.md`](references/technique-index.md)：动画技术索引
- [`references/image-carousel.md`](references/image-carousel.md)：图片轮播技术细节
- [`scripts/build-image-carousel.ps1`](scripts/build-image-carousel.ps1)：本机 PowerPoint 生成脚本

## 核心制作规则

1. 必须提供 5 张图片。
2. 复制整页幻灯片，不要把主栏中的对象全选后粘贴到另一张已有幻灯片中。
3. 这样可以保留前后页面中图片框的一一对应关系，让「平滑」切换正常工作。
4. 同一页内的图片按照逆时针方向填充。
5. 同一张图片在不同页面之间按照顺时针方向进入下一个位置。

## 兼容性

PowerPoint 是主要制作和播放环境。WPS 对「平滑」切换、自动换页和循环播放的支持可能存在差异，正式使用前请在目标软件中预览。

## 来源

本 Skill 根据用户提供的 PPT 图片轮播教学视频整理而成。仓库不包含教程作者的原始图片素材，生成时使用用户自己的图片或其他已获授权的图片。
