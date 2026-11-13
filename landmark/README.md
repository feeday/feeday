# 地标原图 1969 张：Qwen3.5-9B 一次性分类到 Excel

本脚本复用此前本地 Qwen/KoboldCpp 的 OpenAI 兼容多模态接口请求格式，**原图每张只做一次成功的视觉分析**，不修改原始文件。

## 安装及启动（Windows）

1. 确保已有 Qwen3.5-9B 视觉服务在运行，且正确加载视觉投影（mmproj）。
2. 安装依赖：`py -m pip install Pillow`
3. 进入脚本所在目录运行：

先试 20 张：

```bat
py "地标原图_一次分类到Excel_Qwen3.5-9B.py" --input "G:\你的地标目录\Original" --limit 20
```

检查 Excel 中的「详细描述 / 地标身份特征 / 季节 / 天气」后，运行全量：

```bat
py "地标原图_一次分类到Excel_Qwen3.5-9B.py" --input "G:\你的地标目录\Original"
```

也可以不传 `--input`，程序会在控制台询问原图目录。多个输入目录可重复传 `--input`。

## 配置

默认沿用旧脚本中的 `http://10.10.200.59:5002/v1/chat/completions` 以及 `model=koboldcpp`。服务地址改变时：

```bat
py "地标原图_一次分类到Excel_Qwen3.5-9B.py" --input "G:\你的地标目录\Original" --api-url "http://127.0.0.1:5002/v1/chat/completions"
```

默认把图片缩放到最长边 1024 像素、单张请求超时 300 秒、失败最多重试 3 次、每处理 25 张更新一次 Excel。测速时可以尝试 `--max-edge 768`。后端明确支持 `chat_template_kwargs` 时可以加 `--disable-thinking`。

## 表格内容

第一列原始**文件名**，第二列**详细描述**，第三列**稳定的地标身份特征**，随后为地标名称/类型、场景、季节、天气、时间段、拍摄角度/距离、光照、色温/色调、氛围、图像风格、地面状态、构图、天空、植被、周围环境、地标可见程度，末尾附路径、状态、耗时和错误信息。

输出目录默认 `地标原图分类_输出`，内含：

- `地标原图_1969张_详细描述与属性分类.xlsx`：唯一最终表格
- `分类断点缓存.jsonl`：每张成功后立即持久化，作为断点数据
- `失败日志.jsonl`：错误记录；下次运行可重试

重新执行相同命令会跳过文件大小和修改时间均未变化的成功结果。**不要删除断点缓存**，它用于防止重复调用模型。原图不会被移动、覆盖或删除。

注意：不确定季节、天气和时间时填「无法判断」，后续比对生成图时应重点使用「地标身份特征」，不能要求天气、季节、背景必须与原图一致。

---

## Flux / Qwen 生成地标图与 Original 原图双图筛选

新增脚本：[地标生成图_对照原图筛选_Qwen3.5-9B.py](./地标生成图_对照原图筛选_Qwen3.5-9B.py)

原图 1969 张、Flux 816 张、Qwen 1309 张是本项目预计数量；实际按指定目录扫描统计，不强制写死数量。模型通过 OpenAI 兼容多模态接口，一次发送**原图和生成图两张图片**，重点比较地标建筑轮廓、结构、组成关系，不要求天气、时间、构图一致。

Windows 命令（先把示例目录替换成实际路径）：

```bat
py -m pip install Pillow
py "地标生成图_对照原图筛选_Qwen3.5-9B.py" --original "G:\地标\Original" --flux "G:\地标\Flux" --qwen "G:\地标\Qwen" --dry-run --limit 20
py "地标生成图_对照原图筛选_Qwen3.5-9B.py" --original "G:\地标\Original" --flux "G:\地标\Flux" --qwen "G:\地标\Qwen" --limit 20
py "地标生成图_对照原图筛选_Qwen3.5-9B.py" --original "G:\地标\Original" --flux "G:\地标\Flux" --qwen "G:\地标\Qwen"
```

如果已运行上面的 1969 张地标原图分类，可附带：

```bat
py "地标生成图_对照原图筛选_Qwen3.5-9B.py" --original "G:\地标\Original" --flux "G:\地标\Flux" --qwen "G:\地标\Qwen" --reference-xlsx "G:\你的分类结果\地标原图_1969张_详细描述与属性分类.xlsx"
```

匹配优先级：

1. 清理 Original/Flux/Qwen/Edit 等后缀后，按规范化完整文件名匹配
2. 生成图文件名前 5 位原图编号唯一时匹配
3. 没有唯一原图直接 REVIEW；如果这些图只按地标名称分类、没有原图编号，可以手动添加 **--allow-group** 以同地标组的第一张参考图进行建筑身份判定。注意这是“同地标”不是“同一张原图”，表格会明确记录匹配方法

筛选规则：PASS=同一地标主体及结构；FAIL=明确缺失或已换成其他地标；REVIEW=没有可靠原图、遮挡模糊或证据不充分。不会仅因天气、季节、光照、视角、背景改变而FAIL。只有图像本身可见的结构支持同一地标时才能PASS。

默认配置与之前 KoboldCpp 脚本一致：`http://10.10.200.59:5002/v1/chat/completions`、`model=koboldcpp`。默认每张图片最长边768、max_tokens=450、timeout=300秒、每25张刷新总表。地标结构小可尝试 `--max-edge 1024`。后端确认支持时才使用 `--disable-thinking`。

输出目录默认是脚本所在位置的 `地标筛选_输出`：

```text
地标筛选_输出/
├─ PASS/Flux、PASS/Qwen
├─ FAIL/Flux、FAIL/Qwen
├─ REVIEW/Flux、REVIEW/Qwen
├─ 地标生成图_筛选总表.xlsx
├─ 地标生成图_筛选总表.csv
└─ 筛选断点缓存.jsonl
```

只复制，不删除、不移动输入文件；断点缓存防止重复分析。API失败显示ERROR，下一次运行会再次尝试。Excel含生成文件、对应原图、结果、结构判断、文字依据和匹配方式。
