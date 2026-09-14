# PPT Animation Coach

这是一个持续学习、复现和沉淀 PowerPoint 动画技巧的 Codex Skill 项目。

项目不会停留在某一个演示文件。我们会从实际教程视频和用户需求出发，先理解动画逻辑，再整理素材、制作演示、验证效果，最后把已经通过测试的方法加入可复用技术库。

## 统一制作流程

仓库内现有和未来新增的所有动画技术，都遵守同一套流程：

1. **视频切片**

   学习新教程时，先确定研究区间，并按每秒一张图片导出连续切片。用户给出的时间点不需要卡得过死，可以向前、向后多观察约 10 秒。

2. **语言拆解**

   在制作 PPT 前，完整描述动画效果、幻灯片状态、素材数量、对象操作、图层顺序、切换方式和递推规律。说明必须精确到每个素材进行了什么变化。

3. **素材准备与确认**

   列出制作所需的全部素材，检查用户已有文件，并说明还需要寻找或生成什么。完成抠图、裁剪、柔化、虚化等处理后，先把所有候选素材展示给用户。用户确认通过后再制作 PPT。

4. **动画制作与验证**

   按确认后的逻辑制作演示，导出静态和动态预览，与教程对照并修正。需要 PowerPoint「平滑」切换时，必须复制整张幻灯片，保留对象的一一对应关系。

除非用户明确要求跳过素材预览，否则不得在素材确认前直接制作动画。

## 已实现的三项动画技术

### 1. 图片轮播与平滑切换

用户提供五张图片后，可以生成由一个中心大图、两张中图和两张外侧小图组成的轮播。

- 图片沿固定方向连续轮换
- 中心图片始终位于主要视觉层
- 每页由上一页整体复制
- 使用 PowerPoint「平滑」切换
- 图片框保持可编辑

参考：[图片轮播技术说明](references/image-carousel.md)

### 2. 三条文字递进与平滑切换

三条文本依次显现，已有文本同时调整位置。全部文本框从第一页开始就存在，后续只改变位置和透明度。

- 适合三步流程、三段观点和三组关键词
- 文本框保持一一对应
- 不删除或重新创建文本对象
- 文本内容可以继续编辑

参考：[三条文字递进技术说明](references/text-morph-3-line.md)

### 3. 分层视差与平滑切换

两个人物状态完成画面交接，同一主题物体分成远景、中景和前景三个层级。

- 默认使用 9 个清晰远景物体
- 6 个轻微柔化的中景物体
- 3 个明显失焦的前景物体
- 远景移动最少，前景移动最多
- 前景使用预乘透明度虚化，避免白边和暗色矩形
- 两个人物从第一页开始就同时存在，未登场人物放在画面外

调用这项技术时，Skill 会先要求或协助准备：

- 人物状态 A
- 人物状态 B
- 一个清晰的主题物体
- 中景柔化版本
- 前景失焦版本
- 可选背景、标题和副标题

所有素材会先交给用户过目。确认后才会制作动画。

参考：[分层视差技术说明](references/parallax-morph.md)

## 独立能力：从新视频开发动画技术

仓库还包含一项独立的技术开发能力，用来继续学习新的 PPT 动画：

1. 对指定视频区间每秒切一张图
2. 按时间顺序拆解全部操作
3. 用语言说明动画逻辑
4. 列出并准备全部素材
5. 展示素材并等待用户确认
6. 制作和验证演示
7. 用户验收后加入正式技术库

参考：[从新视频学习 PPT 动画](references/learn-animation-from-video.md)

辅助工具：[每秒视频切片脚本](scripts/extract-video-frames.py)

## 演示文件

| 技术 | PowerPoint | 预览 |
| --- | --- | --- |
| 图片轮播 | [下载 PPT](examples/image-carousel-demo.pptx) | [静态预览](examples/image-carousel-preview.png) |
| 三条文字递进 | [下载 PPT](examples/text-morph-3-line-demo.pptx) | [动态预览](examples/text-morph-3-line-preview.mp4) |
| 李白酒杯视差 | [下载 PPT](examples/parallax-morph-li-bai-demo.pptx) | [动态预览](examples/parallax-morph-li-bai-preview.mp4) |

## 项目结构

- [`SKILL.md`](SKILL.md)：Skill 入口、技术路由和统一素材审批流程
- [`references/technique-index.md`](references/technique-index.md)：技术索引
- [`references/learn-animation-from-video.md`](references/learn-animation-from-video.md)：新动画开发流程
- [`scripts/extract-video-frames.py`](scripts/extract-video-frames.py)：每秒切片
- [`scripts/prepare-depth-assets.py`](scripts/prepare-depth-assets.py)：生成中景和前景虚化素材
- [`scripts/build-parallax-morph.ps1`](scripts/build-parallax-morph.ps1)：生成分层视差演示
- [`examples/`](examples/)：已经验证的示例文件

## 兼容性

Microsoft PowerPoint 是主要制作和播放环境。WPS 对「平滑」切换、对象对应、自动换页和循环播放的支持可能存在差异，正式使用前应在目标软件中预览。

## 后续扩展

项目会继续学习更多教程中的文字、图片、视差、镜头运动和组合动画。每项新技术只有在完成切片分析、语言拆解、素材确认和演示验收后，才会加入正式技术索引。
