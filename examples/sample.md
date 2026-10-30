# Markdown Office 导出示例

这是一段带有 **粗体**、*斜体*、`行内代码` 和 [链接](https://example.com) 的正文。

## 普通表格

| 姓名 | 部门 | 分数 |
|:---|:---|---:|
| 张三 | 研发 | 95 |
| 李四 | 产品 | 88 |

## 合并表格

标准 Markdown 管道表格不支持合并单元格，因此使用内嵌 HTML 的
`rowspan` 与 `colspan`：

<table>
  <thead>
    <tr>
      <th rowspan="2">部门</th>
      <th colspan="2">季度</th>
    </tr>
    <tr>
      <th>Q1</th>
      <th>Q2</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td rowspan="2">研发</td>
      <td>10</td>
      <td>12</td>
    </tr>
    <tr>
      <td colspan="2">合计 22</td>
    </tr>
  </tbody>
</table>

## 列表与代码

- 第一项
- 第二项

1. 第一步
2. 第二步

> 这是一段引用内容。

```python
def hello(name: str) -> str:
    return f"Hello, {name}"
```
