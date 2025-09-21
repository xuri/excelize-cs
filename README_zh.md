# excelize-cs

<p align="center"><img width="500" src="https://github.com/xuri/excelize-cs/raw/main/excelize-cs.svg" alt="excelize-cs logo"></p>

<p align="center">
    <a href="https://www.nuget.org/packages/ExcelizeCs"><img src="https://img.shields.io/nuget/v/ExcelizeCs.svg" alt="NuGet version"></a>
    <a href="https://github.com/xuri/excelize-cs/actions/workflows/build.yml"><img src="https://github.com/xuri/excelize-cs/actions/workflows/build.yml/badge.svg" alt="Build Status"></a>
    <a href="https://codecov.io/gh/xuri/excelize-cs"><img src="https://codecov.io/gh/xuri/excelize-cs/branch/main/graph/badge.svg" alt="Code Coverage"></a>
    <a href="https://opensource.org/licenses/BSD-3-Clause"><img src="https://img.shields.io/badge/license-bsd-orange.svg" alt="Licenses"></a>
    <a href="https://www.paypal.com/paypalme/xuri"><img src="https://img.shields.io/badge/Donate-PayPal-green.svg" alt="Donate"></a>
</p>

excelize-cs 是 Go 语言 [Excelize](https://github.com/xuri/excelize) 基础库的 C# 实现，可用于操作 Office Excel 文档，基于 ECMA-376，ISO/IEC 29500 国际标准。可以使用它来读取、写入由 Microsoft Excel&trade; 2007 及以上版本创建的电子表格文档。支持 XLAM / XLSM / XLSX / XLTM / XLTX 等多种文档格式，高度兼容带有样式、图片(表)、透视表、切片器等复杂组件的文档。可应用于各类报表平台、云计算、边缘计算等系统。使用本软件包要求使用的 .Net 版本为 6 或更高版本，C# 语言的版本为 10 或更高版本，获取更多信息请访问 [参考文档](https://xuri.me/excelize/)。

## 快速上手

### 安装

```bash
dotnet add package ExcelizeCs --version 0.0.1
```

### 创建 Excel 文档

下面是一个创建 Excel 文档的简单例子：

```csharp
using ExcelizeCs;

class Program
{
    static void Main()
    {
        var f = Excelize.NewFile();
        try
        {
            // 新建一张工作表
            var index = f.NewSheet("Sheet2");
            // 设置单元格的值
            f.SetCellValue("Sheet2", "A1", "Hello world.");
            f.SetCellValue("Sheet1", "B2", 100);
            // 设置工作簿的默认工作表
            f.SetActiveSheet(index);
            // 根据指定路径保存工作簿
            f.SaveAs("Book1.xlsx");
        }
        catch (RuntimeError err)
        {
            Console.WriteLine(err.Message);
        }
        finally
        {
            var err = f.Close();
            if (!string.IsNullOrEmpty(err))
                Console.WriteLine(err);
        }
    }
}
```

### 读取 Excel 文档

下面是读取 Excel 文档的例子：

```csharp
using ExcelizeCs;

class Program
{
    static void Main()
    {
        ExcelizeCs.File? f;
        try
        {
            f = Excelize.OpenFile("Book1.xlsx");
        }
        catch (RuntimeError err)
        {
            Console.WriteLine(err.Message);
            return;
        }
        try
        {
            // 获取工作表中指定单元格的值
            var cell = f.GetCellValue("Sheet1", "B2");
            Console.WriteLine(cell);
            // 获取 Sheet1 上所有单元格
            var rows = f.GetRows("Sheet1");
            foreach (var row in rows)
            {
                foreach (var c in row)
                {
                    Console.Write($"{c}\t");
                }
                Console.WriteLine();
            }
        }
        catch (RuntimeError err)
        {
            Console.WriteLine(err.Message);
        }
        finally
        {
            // 关闭工作簿
            var err = f.Close();
            if (!string.IsNullOrEmpty(err))
                Console.WriteLine(err);
        }
    }
}
```

### 在 Excel 文档中创建图表

使用 Excelize 生成图表十分简单，仅需几行代码。您可以根据工作表中的已有数据构建图表，或向工作表中添加数据并创建图表。

<p align="center"><img width="650" src="https://github.com/xuri/excelize-cs/raw/main/chart.png" alt="使用 excelize-cs 在 Excel 电子表格文档中创建图表"></p>

```csharp
using ExcelizeCs;

class Program
{
    static void Main()
    {
        var f = Excelize.NewFile();
        var data = new List<List<object?>>
        {
            new() { null, "Apple", "Orange", "Pear" },
            new() { "Small", 2, 3, 3 },
            new() { "Normal", 5, 2, 4 },
            new() { "Large", 6, 7, 8 },
        };
        var text = "Fruit 3D Clustered Column Chart";
        try
        {
            foreach (var row in data)
            {
                f.SetSheetRow(
                    "Sheet1",
                    Excelize.CoordinatesToCellName(1, data.IndexOf(row) + 1),
                    row
                );
            }
            var chart = new Chart
            {
                Type = ChartType.Col3DClustered,
                Series = new ChartSeries[]
                {
                    new()
                    {
                        Name = "Sheet1!$A$2",
                        Categories = "Sheet1!$B$1:$D$1",
                        Values = "Sheet1!$B$2:$D$2",
                    },
                    new()
                    {
                        Name = "Sheet1!$A$3",
                        Categories = "Sheet1!$B$1:$D$1",
                        Values = "Sheet1!$B$3:$D$3",
                    },
                    new()
                    {
                        Name = "Sheet1!$A$4",
                        Categories = "Sheet1!$B$1:$D$1",
                        Values = "Sheet1!$B$4:$D$4",
                    },
                },
                Title = new RichTextRun[] { new() { Text = text } },
            };
            f.AddChart("Sheet1", "E1", chart);
            // 根据指定路径保存文件
            f.SaveAs("Book1.xlsx");
        }
        catch (RuntimeError err)
        {
            Console.WriteLine(err.Message);
        }
        finally
        {
            var err = f.Close();
            if (!string.IsNullOrEmpty(err))
                Console.WriteLine(err);
        }
    }
}
```

### 向 Excel 文档中插入图片

```csharp
using ExcelizeCs;

class Program
{
    static void Main()
    {
        ExcelizeCs.File? f;
        try
        {
            f = Excelize.OpenFile("Book1.xlsx");
        }
        catch (RuntimeError err)
        {
            Console.WriteLine(err.Message);
            return;
        }
        try
        {
            // 插入图片
            f.AddPicture("Sheet1", "A1", "image.png", null);
            // 在工作表中插入图片，并设置图片的缩放比例
            f.AddPicture(
                "Sheet1",
                "A1",
                "image.jpg",
                new GraphicOptions
                {
                    ScaleX = 0.1,
                    ScaleY = 0.1,
                }
            );
            // 在工作表中插入图片，并设置图片的打印属性
            f.AddPicture(
                "Sheet1",
                "A1",
                "image.jpg",
                new GraphicOptions
                {
                    PrintObject = true,
                    LockAspectRatio = false,
                    OffsetX = 15,
                    OffsetY = 10,
                    Locked = false,
                }
            );
            // 保存工作簿
            f.Save();
        }
        catch (RuntimeError err)
        {
            Console.WriteLine(err.Message);
        }
        finally
        {
            // Close the spreadsheet.
            var err = f.Close();
            if (!string.IsNullOrEmpty(err))
                Console.WriteLine(err);
        }
    }
}
```

## 社区合作

欢迎您为此项目贡献代码，提出建议或问题、修复 Bug 以及参与讨论对新功能的想法。

## 开源许可

本项目遵循 BSD 3-Clause 开源许可协议，访问 [https://opensource.org/licenses/BSD-3-Clause](https://opensource.org/licenses/BSD-3-Clause) 查看许可协议文件。

Excel 徽标是 [Microsoft Corporation](https://aka.ms/trademarks-usage) 的商标，项目的图片是一种改编。

gopher.{ai,svg,png} 由 [Takuya Ueda](https://x.com/tenntenn) 创作，遵循 [Creative Commons 3.0 Attributions license](http://creativecommons.org/licenses/by/3.0/) 创作共用授权条款。
