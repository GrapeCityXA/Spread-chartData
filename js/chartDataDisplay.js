window.onload = function () {
    //获取表格
    var spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"), {sheetCount: 3});
    //设置柱形图
    initSpread(spread);
};

//设置柱形图颜色
var colorArray = ['rgb(120, 180, 240)', 'rgb(240, 160, 80)', 'rgb(140, 240, 120)', 'rgb(120, 150, 190)'];

//设置柱形图
function initSpread(spread) {
    var chartType = [{
        //指定chartType为柱形图
        type: GC.Spread.Sheets.Charts.ChartType.columnClustered,
        desc: "columnClustered",
        //设置表格数据
        dataArray: [
            ["", 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
            ["Tokyo", 49.9, 71.5, 106.4, 129.2, 144.0, 176.0, 135.6, 148.5, 216.4, 194.1, 95.6, 54.4],
            ["New York", 83.6, 78.8, 98.5, 93.4, 106.0, 84.5, 105.0, 104.3, 91.2, 83.5, 106.6, 92.3],
            ["London", 48.9, 38.8, 39.3, 41.4, 47.0, 48.3, 59.0, 59.6, 52.4, 65.2, 59.3, 51.2],
            ["Berlin", 42.4, 33.2, 34.5, 39.7, 52.6, 75.5, 57.4, 60.4, 47.6, 39.1, 46.8, 51.1]
        ],
        //设置表格数据展示的位置
        dataFormula: "A1:M5",
        changeStyle: function (chart) {
            //改变文章标题的方法
            changeChartTitle(chart, "The Average Monthly Rainfall");
            //显示数据标签的方法
            changColumnChartDataLabels(chart);
            chart.axes({primaryValue: {title: {text: "Rainfall(mm)"}}});
            //设置柱形图的颜色
            changeChartSeriesColor(chart);
            //设置柱形图的大小和宽度
            changeChartSeriesGapWidthAndOverLap(chart);
        }
    }, {
        //指定chartType为堆积图
        type: GC.Spread.Sheets.Charts.ChartType.columnStacked,
        desc: "columnStacked",
        dataArray: [
            ["", 'Tokyo', 'New York', 'London', 'Berlin'],
            ["The First Quarter", 227.8, 260.9, 127, 110.1],
            ["The Second Quarter", 449.2, 283.9, 136.7, 167.8],
            ["The Third Quarter", 500.5, 300.5, 171, 165.4],
            ["The Fourth Quarter", 344.1, 282.4, 175.7, 137]
        ],
        dataFormula: "A1:E5",
        changeStyle: function (chart) {
            changeChartTitle(chart, "The Average Quarterly Rainfall");
            changColumnChartDataLabels(chart);
            chart.axes({primaryValue: {title: {text: "Rainfall(mm)"}}});
            changeChartSeriesColor(chart);
            changeChartSeriesGapWidthAndOverLap(chart);
        }
    }, {
        //指定chartType为百分比堆积图
        type: GC.Spread.Sheets.Charts.ChartType.columnStacked100,
        desc: "columnStacked100",
        dataArray: [
            ["", 'Tokyo', 'New York', 'London', 'Berlin'],
            ["The First Quarter", 227.8, 260.9, 127, 110.1],
            ["The Second Quarter", 449.2, 283.9, 136.7, 167.8],
            ["The Third Quarter", 500.5, 300.5, 171, 165.4],
            ["The Fourth Quarter", 344.1, 282.4, 175.7, 137]
        ],
        dataFormula: "A1:E5",
        changeStyle: function (chart) {
            changeChartTitle(chart, "The Average Quarterly Rainfall");
            changColumnChartDataLabels(chart);
            chart.axes({primaryValue: {title: {text: "Rainfall(%)"}}});
            changeChartSeriesColor(chart);
            changeChartSeriesGapWidthAndOverLap(chart);
        }
    }];
    var sheets = spread.sheets;
    //挂起活动表单和标签条的绘制
    spread.suspendPaint();
    for (var i = 0; i < chartType.length; i++) {
        var sheet = sheets[i];
        initSheet(sheet, chartType[i].desc, chartType[i].dataArray);
        var chart = addChart(sheet, chartType[i].type, chartType[i].dataFormula);//add chart
        chartType[i].changeStyle(chart);
    }
    //恢复活动表单和标签条的绘制
    spread.resumePaint();
}

//生成柱形图表格的方法（sheetNname是工作表名称，dataArray是工作表的数据）
function initSheet(sheet, sheetName, dataArray) {
    sheet.name(sheetName);
    //为柱形图准备数据
    sheet.setArray(0, 0, dataArray);
    sheet.setColumnWidth(0, 150);
}

//创建图表的方法
function addChart(sheet, chartType, dataFormula) {
    //创建图表
    return sheet.charts.add((sheet.name() + 'Chart1'), chartType, 30, 100, 900, 400, dataFormula, GC.Spread.Sheets.Charts.RowCol.rows);
}

function changeChartTitle(chart, title) {
    chart.title({text: title});
}

//显示数据标签
function changColumnChartDataLabels(chart) {
    var dataLabels = chart.dataLabels();
    dataLabels.showValue = true;
    dataLabels.showSeriesName = false;
    dataLabels.showCategoryName = false;
    var dataLabelPosition = GC.Spread.Sheets.Charts.DataLabelPosition;
    dataLabels.position = dataLabelPosition.outsideEnd;
    chart.dataLabels(dataLabels);
}

//设置柱形图的颜色
function changeChartSeriesColor(chart) {
    var series = chart.series().get();
    for (var i = 0; i < series.length; i++) {
        chart.series().set(i, {backColor: colorArray[i]});
    }
}

//设置柱形图的大小和宽度
function changeChartSeriesGapWidthAndOverLap(chart) {
    var seriesItem = chart.series().get(0);
    seriesItem.gapWidth = 2;
    seriesItem.overlap = 0.1;
    chart.series().set(0, seriesItem);
}