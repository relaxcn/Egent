import OpenAI from "openai";

/**
 * Excel操作的OpenAI Function定义
 * 包含详细的参数描述、示例和最佳实践
 * 确保LLM能够正确理解和调用Excel操作
 */

// =============================================================================
// Excel Function定义常量
// =============================================================================

/**
 * 插入数据到Excel的Function定义
 */
export const INSERT_DATA_FUNCTION: OpenAI.ChatCompletionTool = {
  type: "function",
  function: {
    name: "insertData",
    description: `将数据插入到Excel的指定范围。支持插入各种类型的数据，包括文本、数字、日期等。

使用场景：
- 插入表格数据（如销售记录、统计数据）
- 创建数据列表（如商品清单、员工信息）
- 批量填充单元格内容

最佳实践：
- 确保数据格式统一（每行数据列数相同）
- 对于大量数据，建议分批插入
- 包含标题行时设置includeHeaders为true`,

    parameters: {
      type: "object",
      properties: {
        address: {
          type: "string",
          description: `目标地址，支持多种格式：
- 单个单元格：'A1', 'B2', 'AA10'
- 单元格范围：'A1:C3', 'B2:D5'
- 整列：'A:A', 'B:D'（从第1行开始）
- 整行：'1:1', '3:5'（跨所有列）

示例：'A1:C10' 表示从A1到C10的3列10行区域`,
        },
        data: {
          type: "array",
          description: `要插入的数据，必须是二维数组格式。每个子数组代表一行数据。

数据类型支持：
- 字符串：普通文本、日期格式文本
- 数字：整数、小数
- 布尔值：true/false
- 公式：以'='开头的公式字符串
- null/undefined：空单元格

数据示例：
[
  ["姓名", "年龄", "部门"],           // 标题行
  ["张三", 28, "销售部"],            // 数据行1
  ["李四", 32, "技术部"],            // 数据行2
  ["王五", 25, "市场部"]             // 数据行3
]`,
          items: {
            type: "array",
            description: "单行数据，数组中的每个元素对应一列的值",
            items: {
              type: ["string", "number", "boolean", "null"],
              description: "单元格的值，可以是文本、数字、布尔值或null（空值）"
            }
          }
        },
        includeHeaders: {
          type: "boolean",
          description: `是否包含标题行。默认为false。

- true: 数据的第一行将被视为标题行，可能应用特殊格式（如加粗）
- false: 所有行都被视为普通数据行

建议在插入表格数据时设置为true`,
          default: false
        },
        autoFitColumns: {
          type: "boolean",
          description: `是否自动调整列宽以适应内容。默认为true。

- true: 插入数据后自动调整列宽，确保内容完全可见
- false: 保持原有列宽，内容可能被截断

对于包含较长文本的数据，建议保持为true`,
          default: true
        },
        autoFitRows: {
          type: "boolean",
          description: `是否自动调整行高以适应内容。默认为false。

- true: 插入数据后自动调整行高，适用于多行文本
- false: 保持原有行高

仅在插入包含换行符或长文本的数据时建议设置为true`,
          default: false
        },
        selectAfterInsert: {
          type: "boolean",
          description: `插入数据后是否选中插入的范围。默认为false。

- true: 数据插入后会选中整个插入区域，便于用户查看结果
- false: 保持原有选中状态

在需要用户立即查看插入结果时建议设置为true`,
          default: false
        }
      },
      required: ["address", "data"],
      additionalProperties: false
    }
  }
};

/**
 * 设置单元格背景颜色的Function定义
 */
export const SET_CELL_COLOR_FUNCTION: OpenAI.ChatCompletionTool = {
  type: "function",
  function: {
    name: "setCellColor",
    description: `设置Excel单元格或范围的背景颜色。支持多种颜色格式，可用于突出显示重要数据、分类标记等。

使用场景：
- 突出显示重要数据（如总计行、异常值）
- 按类别给数据着色（如状态指示、优先级标记）
- 创建颜色编码系统（如红色表示警告、绿色表示正常）

颜色建议：
- 红色系：#FF0000, #FF6B6B, #E74C3C（警告、错误）
- 绿色系：#00FF00, #4ECDC4, #2ECC71（成功、正常）
- 蓝色系：#0000FF, #3498DB, #5DADE2（信息、中性）
- 黄色系：#FFFF00, #F39C12, #F7DC6F（注意、待处理）`,

    parameters: {
      type: "object",
      properties: {
        address: {
          type: "string",
          description: `要设置颜色的单元格地址，支持：
- 单个单元格：'A1', 'B2'
- 连续范围：'A1:C3', 'B2:D5'
- 整列：'A:A', 'B:D'
- 整行：'1:1', '3:5'

示例：
- 'A1' - 设置A1单元格的背景色
- 'A1:A10' - 设置A列第1-10行的背景色
- 'B2:D4' - 设置B2到D4矩形区域的背景色`
        },
        backgroundColor: {
          type: "string",
          description: `背景颜色，支持多种格式：

十六进制格式（推荐）：
- '#FF0000' - 红色
- '#00FF00' - 绿色
- '#0000FF' - 蓝色
- '#FFFF00' - 黄色
- '#FFA500' - 橙色

英文颜色名（Excel支持）：
- 'red', 'green', 'blue', 'yellow', 'orange'
- 'pink', 'purple', 'brown', 'gray', 'black'
- 'white', 'silver', 'gold'

RGB格式：
- 'rgb(255, 0, 0)' - 红色
- 'rgb(0, 255, 0)' - 绿色

建议使用十六进制格式以确保颜色准确性`
        }
      },
      required: ["address", "backgroundColor"],
      additionalProperties: false
    }
  }
};

/**
 * 批量设置多个范围背景颜色的Function定义
 */
export const SET_CELL_COLORS_BATCH_FUNCTION: OpenAI.ChatCompletionTool = {
  type: "function",
  function: {
    name: "setCellColorsBatch",
    description: `批量设置多个Excel范围的背景颜色。一次性为多个不同的单元格区域设置不同的颜色，提高效率。

使用场景：
- 同时设置表格中多个区域的颜色（如标题行、数据行、汇总行）
- 为分类数据创建颜色编码（如正面绿色、负面红色、中性黄色）
- 快速应用预设的表格样式（多个区域不同颜色组合）
- 批量处理大量颜色标记任务

优势：
- 一次操作设置多个范围，避免多次调用
- 自动处理错误，提供详细的成功/失败报告
- 支持混合范围类型（单元格、连续范围、整行整列）`,

    parameters: {
      type: "object",
      properties: {
        colorRanges: {
          type: "array",
          description: "颜色范围数组，每个元素包含range（范围地址）和color（颜色）",
          items: {
            type: "object",
            properties: {
              range: {
                type: "string",
                description: `要设置颜色的单元格地址，支持：
- 单个单元格：'A1', 'B2'
- 连续范围：'A1:C3', 'B2:D5'
- 整列：'A:A', 'B:D'
- 整行：'1:1', '3:5'

示例：
- 'A1' - A1单元格
- 'A1:A10' - A列第1-10行
- 'B2:D4' - B2到D4矩形区域`
              },
              color: {
                type: "string",
                description: `背景颜色，支持格式：

十六进制格式（推荐）：
- '#FF0000' - 红色
- '#00FF00' - 绿色
- '#0000FF' - 蓝色
- '#FFFF00' - 黄色
- '#FFA500' - 橙色

英文颜色名：
- 'red', 'green', 'blue', 'yellow', 'orange'
- 'pink', 'purple', 'brown', 'gray'

常用组合建议：
- 标题行：'#4472C4'（深蓝）
- 正面数据：'#70AD47'（绿色）
- 负面数据：'#E15759'（红色）
- 汇总行：'#F2F2F2'（浅灰）`
              }
            },
            required: ["range", "color"],
            additionalProperties: false
          },
          minItems: 1,
          maxItems: 50
        }
      },
      required: ["colorRanges"],
      additionalProperties: false
    }
  }
};

/**
 * 设置单元格字体格式的Function定义
 */
export const SET_CELL_FONT_FUNCTION: OpenAI.ChatCompletionTool = {
  type: "function",
  function: {
    name: "setCellFont",
    description: `设置Excel单元格的完整格式，包括字体、颜色、大小、样式等。用于美化表格、突出重要信息。

使用场景：
- 设置标题行格式（加粗、增大字号）
- 突出显示关键数据（颜色标记、特殊字体）
- 统一表格样式（字体族、大小规范）
- 创建专业的报表格式

格式化建议：
- 标题：Arial/微软雅黑，14-16pt，加粗
- 正文：Arial/宋体，11-12pt，常规
- 重点：红色或蓝色字体，加粗
- 备注：灰色字体，斜体，较小字号`,

    parameters: {
      type: "object",
      properties: {
        address: {
          type: "string",
          description: `要设置格式的单元格地址。支持单个单元格和范围。

示例：
- 'A1' - 单个标题单元格
- 'A1:D1' - 标题行
- 'B2:B10' - 某一列数据
- 'A1:D10' - 整个表格区域`
        },
        fontColor: {
          type: "string",
          description: `字体颜色，格式与背景颜色相同：
- 十六进制：'#000000'（黑色），'#FF0000'（红色）
- 英文名称：'black', 'red', 'blue', 'green'
- RGB：'rgb(0, 0, 0)'

常用颜色：
- 黑色 '#000000' - 正文
- 深蓝 '#003366' - 标题
- 红色 '#CC0000' - 警告/重要
- 深灰 '#666666' - 次要信息`
        },
        fontSize: {
          type: "number",
          description: `字体大小（磅值），常用尺寸：
- 8-9pt：小字、备注
- 10-11pt：正文内容
- 12-14pt：副标题
- 16-18pt：主标题
- 20pt+：大标题

建议范围：8-72pt`,
          minimum: 6,
          maximum: 72
        },
        fontName: {
          type: "string",
          description: `字体名称，建议使用系统常见字体：

中文环境：
- '微软雅黑' - 现代、清晰
- '宋体' - 传统、正式
- '黑体' - 粗壮、醒目
- 'Arial' - 英文标准

英文环境：
- 'Arial' - 无衬线，清晰
- 'Times New Roman' - 有衬线，正式
- 'Calibri' - 现代，Office默认
- 'Helvetica' - 经典无衬线

建议优先使用Arial或微软雅黑以确保兼容性`
        },
        bold: {
          type: "boolean",
          description: `是否加粗字体：
- true：应用加粗，适用于标题、重点内容
- false：取消加粗，恢复常规粗细

加粗建议：
- 表格标题行：建议加粗
- 汇总行/列：建议加粗
- 重要数据：可选加粗
- 普通数据：不建议加粗`
        },
        italic: {
          type: "boolean",
          description: `是否应用斜体：
- true：应用斜体，适用于备注、引用
- false：取消斜体，恢复正常样式

斜体适用场景：
- 备注说明
- 引用内容
- 次要信息
- 特殊标记`
        },
        underline: {
          type: "boolean",
          description: `是否添加下划线：
- true：添加单下划线
- false：移除下划线

下划线用途：
- 超链接样式
- 强调重点
- 分割线效果
- 特殊标记

注意：避免过度使用下划线，以免影响可读性`
        },
        backgroundColor: {
          type: "string",
          description: `背景颜色（可选），与setCellColor功能相同。
可以在设置字体的同时设置背景色，实现一步到位的格式化。

格式同背景颜色设置，支持十六进制、英文名称等`
        },
        horizontalAlignment: {
          type: "string",
          enum: ["left", "center", "right", "justify"],
          description: `水平对齐方式：
- 'left'：左对齐，适用于文本内容
- 'center'：居中对齐，适用于标题、数字
- 'right'：右对齐，适用于数字、金额
- 'justify'：两端对齐，适用于长文本

建议：
- 文本列：左对齐
- 数字列：右对齐或居中
- 标题：居中对齐`
        },
        verticalAlignment: {
          type: "string",
          enum: ["top", "middle", "bottom"],
          description: `垂直对齐方式：
- 'top'：顶端对齐
- 'middle'：居中对齐（推荐）
- 'bottom'：底端对齐

大多数情况建议使用'middle'以获得最佳视觉效果`
        },
        numberFormat: {
          type: "string",
          description: `数字格式代码，控制数值显示方式：

常用格式：
- '@' - 文本格式
- '0' - 整数
- '0.00' - 两位小数
- '#,##0' - 千位分隔符
- '#,##0.00' - 千位分隔符+两位小数
- '¥#,##0.00' - 货币格式
- '0%' - 百分比
- 'm/d/yyyy' - 日期格式
- 'h:mm AM/PM' - 时间格式

示例：
- 金额列：'¥#,##0.00'
- 百分比：'0.00%'
- 日期：'yyyy-mm-dd'`
        }
      },
      required: ["address"],
      additionalProperties: false
    }
  }
};

/**
 * 创建Excel图表的Function定义
 */
export const CREATE_CHART_FUNCTION: OpenAI.ChatCompletionTool = {
  type: "function",
  function: {
    name: "createChart",
    description: `根据指定数据范围创建Excel图表。支持多种图表类型，用于数据可视化和分析。

使用场景：
- 展示趋势变化（折线图）
- 比较数据大小（柱状图）
- 显示占比关系（饼图）
- 分析数据分布（散点图）

选择建议：
- 时间序列数据：折线图（Line）
- 分类数据比较：柱状图（ColumnClustered）
- 部分与整体：饼图（Pie）
- 两变量关系：散点图（XYScatter）`,

    parameters: {
      type: "object",
      properties: {
        type: {
          type: "string",
          enum: [
            "ColumnClustered", "ColumnStacked", "ColumnStacked100",
            "Line", "LineMarkers", "LineMarkersStacked",
            "Pie", "PieExploded",
            "XYScatter", "XYScatterLines",
            "Area", "AreaStacked",
            "Doughnut", "Radar"
          ],
          description: `图表类型，选择最适合数据的可视化方式：

柱状图系列：
- 'ColumnClustered'：簇状柱图，比较不同类别的数据
- 'ColumnStacked'：堆积柱图，显示部分与整体关系
- 'ColumnStacked100'：百分比堆积柱图，显示占比

折线图系列：
- 'Line'：折线图，显示趋势变化
- 'LineMarkers'：带标记的折线图，突出数据点
- 'LineMarkersStacked'：堆积折线图

饼图系列：
- 'Pie'：饼图，显示占比关系
- 'PieExploded'：爆炸饼图，突出某些扇形

散点图系列：
- 'XYScatter'：散点图，分析两变量关系
- 'XYScatterLines'：带线散点图

面积图系列：
- 'Area'：面积图，显示数量变化
- 'AreaStacked'：堆积面积图

特殊图表：
- 'Doughnut'：环形图，类似饼图但中间为空
- 'Radar'：雷达图，多维度数据比较`
        },
        dataRange: {
          type: "string",
          description: `图表的数据源范围，应该包含所有需要图表化的数据：

格式要求：
- 必须是有效的Excel范围：'A1:C5'
- 应包含标题行（第一行作为系列名称）
- 数据应该是连续的矩形区域

数据组织建议：
- 第一行：系列标题或类别名称
- 第一列：X轴标签或类别
- 其他列：数值数据

示例：
- 'A1:B5'：单系列数据（A列标签，B列数值）
- 'A1:D5'：多系列数据（A列标签，B-D列为不同系列）

注意确保数据范围内没有空行或列，否则可能影响图表显示`
        },
        title: {
          type: "string",
          description: `图表标题，应简洁明了地描述图表内容：

标题建议：
- 明确说明数据内容："2024年季度销售额对比"
- 包含时间范围："2023-2024年月度趋势"
- 指明数据单位："营收（万元）变化趋势"
- 保持简洁：避免超过30个字符

如果不提供标题，图表将显示默认标题或无标题`
        },
        left: {
          type: "number",
          description: `图表左边距（像素），相对于工作表的位置：
- 0：贴近工作表左边
- 建议值：50-200像素
- 确保不遮挡现有数据

可配合其他位置参数精确控制图表位置`
        },
        top: {
          type: "number",
          description: `图表上边距（像素），相对于工作表的位置：
- 0：贴近工作表顶部
- 建议值：50-200像素
- 为标题和菜单栏留出空间

通常配合left参数一起使用`
        },
        width: {
          type: "number",
          description: `图表宽度（像素），控制图表的水平尺寸：
- 建议范围：300-800像素
- 过小：内容可能拥挤
- 过大：可能超出屏幕或打印范围

如果不指定，Excel会自动选择合适的宽度`
        },
        height: {
          type: "number",
          description: `图表高度（像素），控制图表的垂直尺寸：
- 建议范围：200-600像素
- 高度应与宽度保持适当比例
- 考虑标题、图例、标签的空间需求

如果不指定，Excel会自动选择合适的高度`
        },
        showLegend: {
          type: "boolean",
          description: `是否显示图例：
- true：显示图例，适用于多系列数据
- false：隐藏图例，适用于单系列或空间有限时

建议：
- 多系列数据：显示图例以区分不同系列
- 单系列数据：可隐藏图例节省空间
- 系列名称明确：可考虑隐藏图例`
        },
        showDataLabels: {
          type: "boolean",
          description: `是否显示数据标签：
- true：在数据点上显示具体数值
- false：隐藏数据标签，保持图表简洁

建议：
- 数据点较少：可显示数据标签
- 数据点较多：避免显示以防重叠
- 精确数值重要：建议显示`
        },
        xAxisTitle: {
          type: "string",
          description: `X轴标题，说明横轴代表的含义：
- 时间轴："月份"、"年份"、"日期"
- 分类轴："产品类别"、"地区"、"部门"
- 数值轴："数量"、"价格"等

简洁明了，通常不超过10个字符`
        },
        yAxisTitle: {
          type: "string",
          description: `Y轴标题，说明纵轴代表的含义：
- 常见："销售额"、"数量"、"百分比"
- 包含单位："营收（万元）"、"增长率（%）"
- 描述性："客户满意度评分"

建议包含数据单位以便理解`
        }
      },
      required: ["type", "dataRange"],
      additionalProperties: false
    }
  }
};

/**
 * 插入Excel公式的Function定义
 */
export const INSERT_FORMULA_FUNCTION: OpenAI.ChatCompletionTool = {
  type: "function",
  function: {
    name: "insertFormula",
    description: `在指定单元格插入Excel公式，用于自动计算和数据处理。支持所有Excel内置函数。

使用场景：
- 数学运算（求和、平均值、最大值等）
- 逻辑判断（IF条件、比较运算）
- 文本处理（连接、提取、格式化）
- 日期时间计算
- 查找引用（VLOOKUP、INDEX等）

公式建议：
- 使用绝对引用（$A$1）固定单元格
- 使用相对引用（A1）允许复制时调整
- 复杂公式可以分步骤在不同单元格中完成`,

    parameters: {
      type: "object",
      properties: {
        address: {
          type: "string",
          description: `要插入公式的目标单元格地址：
- 单个单元格：'A1', 'B10', 'AA100'
- 通常是一个单元格，因为公式需要特定的计算位置

示例：
- 'D10'：在D10单元格插入汇总公式
- 'F2'：在F2单元格插入计算公式`
        },
        formula: {
          type: "string",
          description: `Excel公式，必须以'='开头，遵循Excel公式语法：

数学函数：
- '=SUM(A1:A10)'：求和
- '=AVERAGE(B1:B5)'：平均值
- '=MAX(C1:C10)'：最大值
- '=MIN(C1:C10)'：最小值
- '=COUNT(D1:D10)'：计数

逻辑函数：
- '=IF(A1>100,"大于100","小于等于100")'：条件判断
- '=AND(A1>0,B1<100)'：逻辑与
- '=OR(A1="是",B1="是")'：逻辑或

文本函数：
- '=CONCATENATE(A1," ",B1)'：文本连接
- '=LEFT(A1,3)'：提取左侧字符
- '=LEN(A1)'：文本长度

日期函数：
- '=TODAY()'：今天日期
- '=YEAR(A1)'：提取年份
- '=DATEDIF(A1,B1,"D")'：日期差

查找函数：
- '=VLOOKUP(A1,Sheet2!A:B,2,FALSE)'：垂直查找
- '=INDEX(A:A,MATCH(D1,B:B,0))'：索引匹配

注意事项：
- 公式必须以'='开头
- 字符串用双引号包围
- 使用英文逗号分隔参数
- 函数名使用英文（Excel会自动转换显示语言）`
        },
        calculateImmediately: {
          type: "boolean",
          description: `是否立即计算公式结果：
- true（默认）：插入公式后立即计算，返回结果值
- false：仅插入公式，不强制计算

建议：
- 需要查看结果：设为true
- 仅插入公式：可设为false
- 涉及大量数据：考虑设为false避免性能问题`,
          default: true
        }
      },
      required: ["address", "formula"],
      additionalProperties: false
    }
  }
};

/**
 * 格式化Excel范围的Function定义（与setCellFont功能相同，提供别名支持）
 */
export const FORMAT_RANGE_FUNCTION: OpenAI.ChatCompletionTool = {
  type: "function",
  function: {
    name: "formatRange",
    description: `对指定范围应用完整的格式设置，功能与setCellFont相同。用于批量格式化多个单元格。

这是setCellFont的别名函数，提供相同的功能：
- 字体设置（颜色、大小、样式）
- 背景颜色
- 对齐方式
- 数字格式

特别适用于：
- 格式化整个表格
- 统一多行数据的样式
- 批量应用主题格式`,

    parameters: {
      type: "object",
      properties: {
        address: {
          type: "string",
          description: `要格式化的范围地址，支持大范围选择：
- 多行数据：'A1:D10'
- 整列：'A:D'
- 整行：'1:5'
- 表格区域：'B2:F20'

范围格式化特别适用于统一表格样式`
        },
        fontColor: {
          type: "string",
          description: "字体颜色，格式同setCellFont"
        },
        fontSize: {
          type: "number",
          description: "字体大小，格式同setCellFont",
          minimum: 6,
          maximum: 72
        },
        fontName: {
          type: "string",
          description: "字体名称，格式同setCellFont"
        },
        bold: {
          type: "boolean",
          description: "是否加粗，格式同setCellFont"
        },
        italic: {
          type: "boolean",
          description: "是否斜体，格式同setCellFont"
        },
        underline: {
          type: "boolean",
          description: "是否下划线，格式同setCellFont"
        },
        backgroundColor: {
          type: "string",
          description: "背景颜色，格式同setCellFont"
        },
        horizontalAlignment: {
          type: "string",
          enum: ["left", "center", "right", "justify"],
          description: "水平对齐，格式同setCellFont"
        },
        verticalAlignment: {
          type: "string",
          enum: ["top", "middle", "bottom"],
          description: "垂直对齐，格式同setCellFont"
        },
        numberFormat: {
          type: "string",
          description: "数字格式，格式同setCellFont"
        }
      },
      required: ["address"],
      additionalProperties: false
    }
  }
};

// =============================================================================
// 导出所有Function定义
// =============================================================================

/**
 * 所有Excel操作的Function定义数组
 * 用于OpenAI Function Calling
 */
export const EXCEL_FUNCTIONS: OpenAI.ChatCompletionTool[] = [
  INSERT_DATA_FUNCTION,
  SET_CELL_COLOR_FUNCTION,
  SET_CELL_COLORS_BATCH_FUNCTION,
  SET_CELL_FONT_FUNCTION,
  CREATE_CHART_FUNCTION,
  INSERT_FORMULA_FUNCTION,
  FORMAT_RANGE_FUNCTION
];

/**
 * Function定义的名称映射，便于查找和引用
 */
export const EXCEL_FUNCTION_NAMES = {
  INSERT_DATA: "insertData",
  SET_CELL_COLOR: "setCellColor",
  SET_CELL_COLORS_BATCH: "setCellColorsBatch",
  SET_CELL_FONT: "setCellFont",
  CREATE_CHART: "createChart",
  INSERT_FORMULA: "insertFormula",
  FORMAT_RANGE: "formatRange"
} as const;

/**
 * 获取指定名称的Function定义
 * @param functionName Function名称
 * @returns Function定义或undefined
 */
export function getExcelFunction(functionName: string): OpenAI.ChatCompletionTool | undefined {
  return EXCEL_FUNCTIONS.find(func => (func as any).function?.name === functionName);
}

/**
 * 验证Function调用参数
 * @param functionName Function名称
 * @param args 参数对象
 * @returns 验证结果
 */
export function validateFunctionCall(
  functionName: string,
  args: Record<string, any>
): { valid: boolean; error?: string } {
  const functionDef = getExcelFunction(functionName);

  if (!functionDef) {
    return { valid: false, error: `未知的函数: ${functionName}` };
  }

  const requiredParams = (functionDef as any).function?.parameters?.required || [];

  for (const param of requiredParams) {
    if (!(param in args)) {
      return { valid: false, error: `缺少必需参数: ${param}` };
    }
  }

  return { valid: true };
}