# Markdown 语法示例

这是一个用于测试 `LumiCanvas` Markdown 预览能力的示例文档。

一级标题（Setext）
===

二级标题（Setext）
---

## 标题层级

### 三级标题
#### 四级标题
##### 五级标题
###### 六级标题

---

## 文本样式

普通文本。

**粗体** / *斜体* / _斜体_ / ~~删除线~~

组合示例：**粗体里有 `行内代码`**，以及 *斜体里有链接 [官网](https://github.com/CeeVeeX/LumiCanvas)*。

---

## 引用

> 一级引用
>> 二级引用
>>> 三级引用

---

## 无序列表

- 一级 A
- 一级 B
  - 二级 B.1
    - 三级 B.1.1
* 星号列表也支持

---

## 有序列表

1. 第一项
2. 第二项
   1. 二级项 2.1
   2. 二级项 2.2
3) 使用右括号分隔的第三项

---

## 任务列表

- [ ] 待办事项
- [x] 已完成事项
  - [ ] 嵌套待办
  - [x] 嵌套完成

---

## 链接

行内链接：[LumiCanvas 仓库](https://github.com/CeeVeeX/LumiCanvas)

自动链接（URL）：https://learn.microsoft.com/dotnet/

自动邮箱链接：example@test.com

---

## 图片

> 将下面地址替换为你本机可访问的图片路径或网络图片地址。

![示例图片](https://raw.githubusercontent.com/github/explore/main/topics/markdown/markdown.png)

[![带链接图片](https://raw.githubusercontent.com/github/explore/main/topics/windows/windows.png)](https://www.microsoft.com/windows)

---

## 分割线

上面和下面都是分割线：

***

___

---

## 表格

| 功能 | 语法 | 说明 |
| --- | --- | --- |
| 粗体 | `**text**` | 行内样式 |
| 斜体 | `*text*` / `_text_` | 行内样式 |
| 删除线 | `~~text~~` | 行内样式 |
| 链接 | `[name](url)` | 可点击 |

---

## 行内代码与代码块

行内代码示例：`var message = "Hello Markdown";`

```csharp
using System;

Console.WriteLine("Hello from fenced code block");
```

~~~sql
SELECT Id, Title
FROM Notes
WHERE IsArchived = 0
ORDER BY UpdatedAt DESC;
~~~

---

## 转义字符

这些字符被转义后会按字面显示：

\*不是斜体\*，\_不是斜体\_，\`不是代码\`

\[不是链接\](https://example.com)

\# 不是标题

---

## 脚注

这是一段带脚注的文本[^note1]，这里还有第二个脚注[^note2]。

[^note1]: 这是第一个脚注内容。
[^note2]: 这是第二个脚注，支持普通文本说明。