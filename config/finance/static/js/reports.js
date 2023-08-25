$(document).ready(function() {
    $('#saveEdit').on('click', function() {
        // var csrfToken = $('input[name="csrfmiddlewaretoken"]').val();
        // var activitiesContent = CKEDITOR.instances.activities_form.getData();
        // var accomplishmentsContent = CKEDITOR.instances.accomplishments_form.getData();
        // console.log(activitiesContent);
        // var formData = {
        //     csrfmiddlewaretoken: csrfToken,
        //     activities: activities,
        //     accomplishments: accomplishments
        // };
        // $.ajax({
        //     url: '/reports_advantage/',  // URL to your view
        //     type: 'POST',
        //     dataType: 'json',
        //     data: formData,
        //     success: function(reponse) {
        //         alert('Changes saved.')
        //     },
        //     error: function(response) {
        //         alert(response);
        //     }
        // });
        $('#myModal').modal('hide');  // Close the modal
        location.reload();  // Refresh the page
    });

    data = [{
        date: new Date(2021, 0, 1).getTime(),
        value: 1000
    }, {
        date: new Date(2021, 0, 2).getTime(),
        value: 800
    }, {
        date: new Date(2021, 0, 3).getTime(),
        value: 900
    }, {
        date: new Date(2021, 0, 4).getTime(),
        value: 800
    }, {
        date: new Date(2021, 0, 5).getTime(),
        value: 800
    }];
    createLineChart('chartdiv', data);
    createLineChart('chartdiv2', data);
    createLineChart('chartdiv3', data);

    function createLineChart(chartId, data) {
        const root = am5.Root.new(chartId);
        root.setThemes([
            am5themes_Animated.new(root)
        ]);
        // instantiating the line chart
        var chart = root.container.children.push(am5xy.XYChart.new(root, {
            panX: true,
            panY: true,
            wheelX: "panX",
            wheelY: "zoomX",
            pinchZoomX: true
        }));

        chart.get("colors").set("step", 3);

        // Add cursor
        // https://www.amcharts.com/docs/v5/charts/xy-chart/cursor/
        var cursor = chart.set("cursor", am5xy.XYCursor.new(root, {}));
        cursor.lineY.set("visible", false);

        // creating the axis for the chart
        var yAxis = chart.yAxes.push(am5xy.ValueAxis.new(root, {
            maxDeviation: 0.3,
            renderer: am5xy.AxisRendererY.new(root, {})
        }));

        var xAxis = chart.xAxes.push(am5xy.DateAxis.new(root, {
            maxDeviation: 0.3,
            baseInterval: {
                timeUnit: "day",
                count: 1
            },
            renderer: am5xy.AxisRendererX.new(root, {}),
            tooltip: am5.Tooltip.new(root, {})
        }));

        var series = chart.series.push(am5xy.LineSeries.new(root, {
            name: "Series 1",
            xAxis: xAxis,
            yAxis: yAxis,
            valueYField: "value",
            valueXField: "date",
            tooltip: am5.Tooltip.new(root, {
                labelText: "{valueY}"
            })
        }));
        series.strokes.template.setAll({
            strokeWidth: 2,
            strokeDasharray: [3, 3]
        });

        series.data.setAll(data);

        // Make stuff animate on load
        // https://www.amcharts.com/docs/v5/concepts/animations/
        series.appear(1000);
        chart.appear(1000, 100);
    }
});

