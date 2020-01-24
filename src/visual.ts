//import { text } from "d3";

/*
 *  Power BI Visual CLI
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

// 'no-any' used at 83 places because return type is object
module powerbi.extensibility.visual {
    'use strict';
    import TooltipEventArgs = powerbi.extensibility.utils.tooltip.TooltipEventArgs;
    import ITooltipServiceWrapper = powerbi.extensibility.utils.tooltip.ITooltipServiceWrapper;
    import createTooltipServiceWrapper = powerbi.extensibility.utils.tooltip.createTooltipServiceWrapper;
    import ValueFormatter = powerbi.extensibility.utils.formatting.valueFormatter;
    import IValueFormatter = powerbi.extensibility.utils.formatting.IValueFormatter;
    import ISelectionId = powerbi.visuals.ISelectionId;
    import TextProperties = powerbi.extensibility.utils.formatting.TextProperties;
    import TextMeasurementService = powerbi.extensibility.utils.formatting.textMeasurementService;
    import createLegend = powerbi.extensibility.utils.chart.legend.createLegend;
    import ILegend = powerbi.extensibility.utils.chart.legend.ILegend;
    import Legend = powerbi.extensibility.utils.chart.legend;
    import LegendData = powerbi.extensibility.utils.chart.legend.LegendData;
    import LegendDataPoint = powerbi.extensibility.utils.chart.legend.LegendDataPoint;
    import LegendIcon = powerbi.extensibility.utils.chart.legend.LegendIcon;
    import LegendPosition = powerbi.extensibility.utils.chart.legend.LegendPosition;
    import DataViewObjectsParser = powerbi.extensibility.utils.dataview.DataViewObjectsParser;
    import DataView = powerbi.DataView;
    import IViewport = powerbi.IViewport;
    interface IDataPoints {
        value: string[];
        index: number;
        selectionId: ISelectionId;
    }

    interface IViewModel {
        dataPoints: IDataPoints[];
        columns: DataViewMetadataColumn[];
    }

    function visualTransform(options: VisualUpdateOptions, host: IVisualHost): IViewModel {
        const dataViews: DataView[] = options.dataViews;
        const viewModel: IViewModel = {
            dataPoints: [],
            columns: []
        };

        const category: DataViewCategoryColumn = dataViews[0].categorical.categories[0];
        const len: number  = dataViews[0].categorical.categories[0].values.length;
        const categoriesLength: number = dataViews[0].categorical.categories.length;
        const valuesLength: number = dataViews[0].categorical.values.length;
        const cLength: number = dataViews[0].metadata.columns.length;
        // tslint:disable-next-line:no-any
        const rows1: any[] = [] ;
        let iRow : number;
        let iColumn : number ;
        let kValue: number = 0;
        let jValue: number = 0;
        for (iRow = 0; iRow < len; iRow++) {
            rows1[iRow] = [];
            kValue = 0 , jValue = 0;
            for (iColumn = 0; iColumn < cLength; iColumn++) {
                if (dataViews[0].metadata.columns[iColumn].isMeasure === true) {
            rows1[iRow][iColumn] = dataViews[0].categorical.values[kValue++].values[iRow] ;
            } else {
                rows1[iRow][iColumn] = dataViews[0].categorical.categories[jValue++].values[iRow];
            }
         }
        }
        // tslint:disable-next-line:no-any
        const rows: any[] = rows1;
        viewModel.columns = dataViews[0].metadata.columns;

        rows.forEach(function (row: DataViewCategorical, index: number): void {

            viewModel.dataPoints.push({
                value: <string[]> row,
                index: index,
                selectionId: host.createSelectionIdBuilder()
                    .withCategory(category, index)
                    .createSelectionId()
            });
        });

        return viewModel;
    }

    export class Visual implements IVisual {
        private showXAxisIsOn: boolean;
        private target: HTMLElement;
        private settings: VisualSettings;
        private selectionManager: ISelectionManager;
        private host: IVisualHost;
        private viewModel: IViewModel;
        private selectionIdBuilder: ISelectionIdBuilder;
        // tslint:disable-next-line:no-any
        private topCont: any;
        // tslint:disable-next-line:no-any
        private mainCont: any;
        // tslint:disable-next-line:no-any
        private xAxisCont: any;
        // tslint:disable-next-line:no-any
        private yAxisCont: any;
        private groupedColumn: number;
        private targetColumn: number;
        private binColumn: number;
        private iColumn: number;
        private jColumn: number;
        private chartType: string;
        private maxValue: number = 0;
        private minValue: number = 0;
        private averageValue: number;
        private numberOfBins: number = 0;
        private $colorByCont: JQuery;
        private tooltipServiceWrapper: ITooltipServiceWrapper;
        private i: boolean = false;
        private colorByValuesLength: number = 0;
        // tslint:disable-next-line:no-any
        private currentViewport: IViewport;
        // tslint:disable-next-line:no-any
        private options: any;
        private legend: ILegend;
        private legendObjectProperties: DataViewObject;
        private iColumntext: number;
        private jColumntext: number;
        private previousLength: number = 0;
        private renderedTime: number = 0;
        private top: number = 40;
        private pixelLiteral : string = 'px';
        //undo variables
        private groupedColumnOld: number[] = [];
        private targetColumnOld: number[] = [];
        private binColumnOld: number[] = [];
        private colorByColumnOld: number[] = [];
        private chartTypeOld: string[] = [];
        private actions: number[] = [];
        private currentDataview: DataView;
        // tslint:disable-next-line:no-any
        private selectionsOld: any[] = [];
        // tslint:disable-next-line:no-any
        private globalSelectionsOld: any[] = [];
        private iColumnold: number[] = [];
        private jColumnOld: number[] = [];
        private iColumntextold: number[] = [];
        private jColumntextOld: number[] = [];
        // tslint:disable-next-line:no-any
        private globalSelections: any = [];
        private margin: number;
        // tslint:disable-next-line:no-any
        private selectionIndexes: any[] = [];
        private dataView: DataView;
        private canUndo: boolean = false;
        private width: number;
        private numberCategory: boolean;
        private maxLineIsOn: boolean;
        private showlegend: boolean;
        private showXAxis: boolean;
        private constantLineIsOn: boolean;
        private constantLineValue: number;
        private showXAxisValue: number;
        private timesUpdateCalled: number;
        private chartRendered: boolean = false;
        private minLineIsOn: boolean;
        private avgLineIsOn: boolean;
        private maxLineStyle: string;
        private minLineStyle: string;
        private avgLineStyle: string;
        private maxLineFill: string;
        private minLineFill: string;
        private avgLineFill: string;
        private maxLineOpacity: number;
        private minLineOpacity: number;
        private avgLineOpacity: number;
        private maxLineDataLabel: boolean;
        private minLineDataLabel: boolean;
        private avgLineDataLabel: boolean;
        private constantLineStyle: string;
        private constantLineFill: string;
        private constantLineOpacity: number;
        private constantLineDataLabel: boolean;
        private value: string;
        private colorByDataSelection: number = 0;
        private colorByColumn: number;
        private previouscolorby: number;
        private $mainCont: JQuery;
        private mainContWidth: number;
        private mainContHeight: number;
        private $binningCont: JQuery;
        private $groupingCont: JQuery;
        private $chartTypeMenu: JQuery;
        private $dynamicBinCont: JQuery;
        private $yAxis: JQuery;
        private $xAxis: JQuery;
        private $targetCont: JQuery;
        private $colorCont: JQuery;
        private $colorContShape: JQuery;
        private $colorContText: JQuery;
        private $columnSelector: JQuery;
        private $xAxisLabel: JQuery;
        private $yAxisLabel: JQuery;
        private flag: number = 0;
        private selectedColumnsString: string = '';
        private isCategory: boolean = true;
        private height: number;
        // tslint:disable-next-line:no-any
        private rootElement: any;
        private previousDataLength: number = 0;
        private previousColorByColumn: number = 0;
        private counter: number = 0;
        private updateCalled: boolean = false;
        // tslint:disable-next-line:no-any
        private barcolor: any = [];
        private showLegend: boolean;
        private noOfBars: number;
        private noOfColumns: number;
        private barwidth: number;
        private labelWidth: number;
        private countOfMenuItems: number;
        // ColumnSelector
        private dataColIndex: number[] = [];
        // Sorting in Grid
        private gridRow: d3.selection.Update<DataViewTableRow>[] = [];
        // tslint:disable-next-line:no-any
        private binData: any[] = [];
        // Arrays :  value for every bin
        private isSorted: boolean[] = [];
        private isSortAsc: boolean[] = [];
        private sortedColumnIndex: number[] = [];
        private space: string = ' ';
        // tslint:disable-next-line:no-any
        private svgElement: any;
        // Sorting column names in dropdown
        private dropdown: d3.selection.Update<DataViewTableRow>[] = [];
        // Isolated points in Grid
        private keyColumnIndex: number;
        private selectedBins: number[] = [];
        private type: boolean = false;
        // Constants
        private dot: string = '.';
        private px: string = 'px';
        private prevChartType: string;
        private undoPressed: boolean = false;
        // tslint:disable-next-line:no-any
        private prevGlobalSelections: any[] = [];
        // Formatters
        //Value Formatter Creation
        private targetColumnformatter: IValueFormatter;
        private binColumnformatter: IValueFormatter;
        private dataPaneColumns: DataViewMetadataColumn[] = [];
        //Last Selected bar
        private lastSelectedbar: number;
        // Add colors to this array
        private colors: string[][] = [['#330000', '#331900', '#333300', '#193300', '#003300',
            '#003319', '#003333', '#001933', '#000033', '#190033', '#330033', '#330019', '#000000'],
        ['#660000', '#663300', '#666600', '#336600', '#006600', '#006633',
            '#006666', '#003366', '#000066', '#330066', '#660066', '#660033', '#202020'],
        ['#990000', '#994C00', '#999900', '#4C9900', '#009900', '#00994C', '#009999',
            '#004C99', '#000099', '#4C0099', '#990099', '#99004C', '#404040'],
        ['#CC0000', '#CC6600', '#CCCC00', '#66CC00', '#00CC00', '#00CC66', '#00CCCC', '#0066CC',
            '#0000CC', '#6600CC', '#CC00CC', '#CC0066', '#606060'],
        ['#FF0000', '#FF8000', '#FFFF00', '#80FF00', '#00FF00', '#00FF80', '#00FFFF',
            '#0080FF', '#0000FF', '#7F00FF', '#FF00FF', '#FF007F', '#808080'],
        ['#FF3333', '#FF9933', '#FFFF33', '#99FF33', '#33FF33', '#33FF99', '#33FFFF', '#3399FF', '#3333FF',
            '#9933FF', '#FF33FF', '#FF3399', '#A0A0A0'],
        ['#FF6666', '#FFB266', '#FFFF66', '#B2FF66', '#66FF66', '#66FFB2', '#66FFFF', '#66B2FF', '#6666FF',
            '#B266FF', '#FF66FF', '#FF66B2', '#C0C0C0'],
        ['#FF9999', '#FFCC99', '#FFFF99', '#CCFF99', '#99FF99', '#99FFCC', '#99FFFF', '#99CCFF', '#9999FF', '#CC99FF', '#FF99FF',
            '#FF99CC', '#E0E0E0'],
        ['#FFCCCC', '#FFE5CC', '#FFFFCC', '#E5FFCC', '#CCFFCC', '#CCFFE5', '#CCFFFF', '#CCE5FF', '#CCCCFF', '#E5CCFF', '#FFCCFF',
            '#FFCCE5', '#FFFFFF']];
        //Display Ratio
        private ratio: number = 20;
        private textSize: number;
        private textWidth: number;
        constructor(options: VisualConstructorOptions) {
            this.target = options.element;
            this.host = options.host;
            this.selectionManager = options.host.createSelectionManager();
            this.options = options;
            this.svgElement = options.element;
            this.tooltipServiceWrapper = createTooltipServiceWrapper(
                this.host.tooltipService,
                options.element);
            this.rootElement = options.element;
            d3.select(this.target).style('cursor', 'default');
            this.topCont = d3.select(this.target).append('div')
                .classed('topCont', true);
            this.mainCont = d3.select(this.target).append('div')
                .classed('mainCont', true);
            this.xAxisCont = d3.select(this.target).append('div')
                .classed('xAxisCont', true);
            this.yAxisCont = d3.select(this.target).append('div')
                .classed('yAxisCont', true);
            this.$mainCont = $('.mainCont');
            this.legend = createLegend(options.element, false, null, true);
        }
        // tslint:disable-next-line:no-any
        private renderYAxis(container: any, contClass: any, labelClass: any, columnClass: any): void {
            const thisObj: this = this;

            /*For Y Axis*/
            // tslint:disable-next-line:no-any
            const yAxis: any = container.append('div').classed(`presentationCont ${contClass}`, true);
            if (contClass === 'groupingCont') {
                thisObj.$groupingCont = $('.groupingCont');
                const menuY: JQuery = $('.menuY');
                yAxis.style('left', menuY.position().left + this.px)
                    .style('top', ((menuY.position().top + thisObj.top) + this.px));
            } else {
                thisObj.$yAxis = $('.yAxis');
            }

            // tslint:disable-next-line:no-any
            const columns: any = yAxis.selectAll('p')
                .data(thisObj.viewModel.columns).enter()
                .append('p')
                .attr({
                    class: function (datum: DataViewMetadataColumn, index: number): string {
                        return `dropdown ${columnClass} ${datum.displayName} index${index}`;
                    },
                    'data-id': function (datum: DataViewMetadataColumn, index: number): number { return index; }
                })
                .text(function (datum: DataViewMetadataColumn): string { return datum.displayName; });

            //Sorting the columns in Y Axis dropdown
            thisObj.dropdown = yAxis.selectAll(`.dropdown`).data(thisObj.viewModel.columns);
            thisObj.columnSorter();

            d3.select(`.${labelClass}`).on('click', function (): void {
                if ('yAxis' === contClass) {
                    thisObj.hideMenus(thisObj.$yAxis);
                } else if ('groupingCont' === contClass) {
                    thisObj.hideMenus(thisObj.$groupingCont);
                }
            });

            columns.on('click', function (): void {
                d3.selectAll(`.${columnClass}`).style('border', '0');
                thisObj.cacheOldMenuState();
                if (thisObj.chartType.toLowerCase() === 'bar') {
                    thisObj.groupedColumn = parseInt(d3.select(this).attr('data-id'), 10);
                } else {
                    thisObj.targetColumn = parseInt(d3.select(this).attr('data-id'), 10);
                }
                thisObj.persistOrient();
                thisObj.setOrient(thisObj.options);
                $(`.${contClass}`).hide();
            });

            $('.menuY p').text(thisObj.viewModel.columns[thisObj.chartType.toLowerCase() === 'bar' ?
                thisObj.groupedColumn : thisObj.targetColumn].displayName)
                .attr('title', thisObj.viewModel.columns[thisObj.chartType.toLowerCase() === 'bar' ?
                    thisObj.groupedColumn : thisObj.targetColumn].displayName);

            const names: string[] = thisObj.viewModel.columns[thisObj.chartType.toLowerCase() === 'bar' ?
                thisObj.groupedColumn : thisObj.targetColumn].displayName.split(' ', thisObj.viewModel.columns.length);
            const jointName: string = names.join('.');
            let column: number;
            if (thisObj.chartType.toLowerCase() === 'bar') {
                column = thisObj.groupedColumn;
            } else {
                column = thisObj.targetColumn;
            }

            d3.select(`.${contClass} .${jointName}.index${column}`).style('background-color', '#4c4c4c');
        }

        // tslint:disable-next-line:no-any
        private renderXAxis(container: any, contClass: string, labelClass: string, columnClass: string): void {
            const thisObj: this = this;
            const menuX: JQuery = $('.menuX');

            // tslint:disable-next-line:no-any
            const xAxis: any = container.append('div').classed(`presentationCont ${contClass}`, true);
            if (contClass === 'targetCont') {
                thisObj.$targetCont = $('.targetCont');
                xAxis.style('left', menuX.position().left + this.px)
                    .style('top', ((menuX.position().top + thisObj.top) + this.px));
            } else {
                thisObj.$xAxis = $('.xAxis');
            }
            // tslint:disable-next-line:no-any
            const columns: any = xAxis.selectAll('p')
                .data(thisObj.viewModel.columns).enter()
                .append('p')
                .attr({
                    class: function (datum: DataViewMetadataColumn, index: number): string {
                        return `dropdown ${columnClass} ${datum.displayName} index${index}`;
                    },
                    'data-id': function (datum: DataViewMetadataColumn, index: number): number { return index; }
                })
                .text(function (datum: DataViewMetadataColumn): string { return datum.displayName; });

            // Sorting the columns in X Axis dropdown
            thisObj.dropdown = xAxis.selectAll(`.dropdown`).data(thisObj.viewModel.columns);
            thisObj.columnSorter();

            d3.select(`.${labelClass}`).on('click', function (): void {
                if ('xAxis' === contClass) {
                    thisObj.hideMenus(thisObj.$xAxis);
                } else if ('targetCont' === contClass) {
                    thisObj.hideMenus(thisObj.$targetCont);
                }
            });

            columns.on('click', function (): void {
                d3.selectAll(`.${columnClass}`).style('border', '0');
                thisObj.cacheOldMenuState();
                if (thisObj.chartType.toLowerCase() === 'bar') {
                    thisObj.targetColumn = parseInt(d3.select(this).attr('data-id'), 10);
                } else {
                    thisObj.groupedColumn = parseInt(d3.select(this).attr('data-id'), 10);
                }
                thisObj.persistOrient();
                thisObj.setOrient(thisObj.options);
                $(`.${contClass}`).hide();
            });
            $('.menuX p')
                .text(thisObj.viewModel.columns[thisObj.chartType.toLowerCase() === 'bar' ?
                    thisObj.targetColumn : thisObj.groupedColumn].displayName)
                .attr('title', thisObj.viewModel.columns[thisObj.chartType.toLowerCase() === 'bar' ?
                    thisObj.targetColumn : thisObj.groupedColumn].displayName);
            const names: string[] = (thisObj.viewModel.columns[thisObj.chartType.toLowerCase() === 'bar' ?
                thisObj.targetColumn : thisObj.groupedColumn].displayName).split(' ', thisObj.viewModel.columns.length);
            const jointName: string = names.join('.'); let column: number;
            if (thisObj.chartType.toLowerCase() === 'bar') {
                column = thisObj.targetColumn;
            } else {
                column = thisObj.groupedColumn;
            }
            d3.select(`.${contClass} .${jointName}.index${column}`).style('background-color', '#4c4c4c');
        }

        // Function to sort columns in dropdowns
        private columnSorter(): void {
            // tslint:disable-next-line:no-any
            this.dropdown.sort(function (value1: any, value2: any): number {
                // tslint:disable-next-line:no-any
                const val1: any = isNaN(+value1.displayName) ?
                    value1.displayName && value1.displayName.toString().toLowerCase() || '' : value1.displayName || '';
                // tslint:disable-next-line:no-any
                const val2: any = isNaN(+value2.displayName) ?
                    value2.displayName && value2.displayName.toString().toLowerCase() || '' : value2.displayName || '';
                const result: number = val1 > val2 ? 1 : -1;

                return result;
            });
        }
        private renderMenu(): void {
            d3.select('.topCont').remove();
            let elem: HTMLElement;
            elem = document.createElement('div');
            elem.className = 'topCont';
            $(this.target).prepend(elem);
            const px: string = 'px';

            this.topCont = d3.select('.topCont');

            // Appending the tabs to menu
            const thisObj: this = this;

            d3.selectAll('.presentationCont').remove();

            // Tabs in Menu Pane
            const topMenuItems: string[] = ['View as', 'Binning by', 'X', 'Y',
                'Color', 'Text Color', 'Bin Size', 'Color By', 'Undo'];
            const menuItemsClassNames: string[] =
                ['ViewAs', 'Binningby', 'X', 'Y', 'Color', 'TextColor', 'RangeforBinning', 'ColorBy', 'Undo'];
            thisObj.countOfMenuItems = menuItemsClassNames.length;
            if (thisObj.chartType === 'Brick') {
                topMenuItems[2] = 'Label';
                topMenuItems[3] = 'Value';
            }

            enum chartTypes { 'Brick', 'Bar', 'Table', 'Column' }

            const chartTypeIcons: string[] = ['brickIcon', 'barIcon', 'gridIcon', 'columnIcon'];
            // Appending the tabs to menu
            // tslint:disable-next-line:no-any
            const topMenu: any = thisObj.topCont.append('ul').classed('presentationCont topMenu', true);
            //here

            topMenu.selectAll('li')
                .data(topMenuItems).enter()
                .append('li')
                .attr({
                    class: function (datum: string, index: number): string { return `menu${menuItemsClassNames[index]} topMenuOptions`; },
                    'data-id': function (datum: string, index: number): number { return index; }
                })
                .style('display', 'table')
                .style('width', `${Math.floor(100 / thisObj.countOfMenuItems)}` + '%')
                .text(function (datum: string): string { return datum; });

            // No drop down icon for last 2 bttons
            for (let iterator: number = 0; iterator < topMenuItems.length - 1; iterator++) {
                topMenu.select(`.menu${menuItemsClassNames[iterator]}.topMenuOptions`)
                    .append('span')
                    .classed('dropdownIcon', true);

            }
            for (let iterator: number = 0; iterator < topMenuItems.length; iterator++) {
                topMenu.select(`.menu${menuItemsClassNames[iterator]}.topMenuOptions`)
                    .append('p')
                    .attr({
                        class: function (datum: string, index: number): string { return `label${index}`; },
                        'data-id': function (datum: string, index: number): number { return index; }
                    });
            }
            //Hide X and Y axis tabs from top menu pane when chart type is table
            if (thisObj.chartType === 'Table') {
                $('li.menuX').hide();
                $('li.menuY').hide();
            } else {
                $('li.menuX').show();
                $('li.menuY').show();
                // Remove column selector when chart type is not table
                $(`.columnSelectorLabel`).remove();
            }
            d3.select('.menuUndo')
                .classed('undoCont', true)
                .on('click', () => {
                    if (thisObj.actions.length === 0) {
                        return;
                    }
                    const action: number = thisObj.actions.pop();
                    if (action === 1) {
                        thisObj.binColumn = thisObj.binColumnOld.pop();
                        thisObj.colorByColumn = thisObj.colorByColumnOld.pop();
                        thisObj.groupedColumn = thisObj.groupedColumnOld.pop();
                        thisObj.targetColumn = thisObj.targetColumnOld.pop();
                        thisObj.chartType = thisObj.chartTypeOld.pop();
                        thisObj.iColumn = thisObj.iColumnold.pop();
                        thisObj.jColumn = thisObj.jColumnOld.pop();
                        thisObj.iColumntext = thisObj.iColumntextold.pop();
                        thisObj.jColumntext = thisObj.jColumntextOld.pop();
                        thisObj.persistOrient();
                        thisObj.setOrient(this.options);
                    } else if (action === 2) {
                        thisObj.selectionIndexes = thisObj.selectionsOld.pop();
                        thisObj.globalSelections = thisObj.globalSelectionsOld.pop();
                        thisObj.undoPressed = true;
                        if (0 === thisObj.selectionsOld.length) {
                            thisObj.canUndo = false;
                           // d3.select('.topCont').style('margin-bottom', '33px');
                        }
                        thisObj.renderChart();
                    }
                })
                .append('span')
                .classed('undo', true)
                .attr('title', 'Undo');
            // Views Menu- Appending Images to Views
            // tslint:disable-next-line:no-any
            const chartTypeCont: any = thisObj.topCont.append('div').classed('presentationCont chartTypeMenu', true);
            thisObj.$chartTypeMenu = $('.chartTypeMenu');

            chartTypeCont.selectAll('.chartTypeMenu')
                .data(chartTypeIcons).enter()
                .append('div')
                .attr({
                    class: function (datum: string, index: number): string { return `imageIcon ${datum} ${chartTypes[index]}`; },
                    'data-id': function (datum: string, index: number): number { return index; },
                    title: function (datum: string, index: number): string { return (`${chartTypes[index]} Chart`); }
                });

            thisObj.$chartTypeMenu.hide();

            d3.select('.menuViewAs').on('click', function (): void {
                thisObj.hideMenus(thisObj.$chartTypeMenu);
            });
            // tslint:disable-next-line:no-any
            const $imageIcon: any = d3.selectAll('.imageIcon');
            d3.select(`.${thisObj.chartType}`).style({ border: '1px solid' });
            $imageIcon.on('click', function (): void {
                let index: number;
                index = parseInt(d3.select(this).attr('data-id'), 10);
                thisObj.chartType = chartTypes[index];
                if (thisObj.prevChartType !== thisObj.chartType) {
                    $imageIcon.style({ border: '0', 'background-color': 'white' });
                    d3.select(this).style({ border: '1px solid' });
                    thisObj.cacheOldMenuState();
                    thisObj.persistOrient();
                    thisObj.setOrient(thisObj.options);
                }
            });
            $('.menuViewAs p').text(thisObj.chartType ? thisObj.chartType : 'None');
            /*For Binning Dropdown*/
            const $menuBinningBy: JQuery = $('.menuBinningby');
            // tslint:disable-next-line:no-any
            const binningCont: any = thisObj.topCont.append('div').classed('presentationCont binningCont', true)
                .style('left', $menuBinningBy.position().left + px)
                .style('top', (($menuBinningBy.position().top + thisObj.top) + px));
            thisObj.$binningCont = $('.binningCont');
            binningCont.selectAll('p')
                .data(thisObj.viewModel.columns).enter()
                .append('p')
                .attr({
                    class: function (datum: DataViewMetadataColumn, index: number): string {
                        return `dropdown binColumn ${datum.displayName} index${index}`;
                    },
                    'data-id': function (datum: DataViewMetadataColumn, index: number): number { return index; }
                })
                .text(function (datum: DataViewMetadataColumn): string { return datum.displayName; });

            //Sorting the columns in binning dropdown
            thisObj.dropdown = binningCont.selectAll(`.dropdown`).data(thisObj.viewModel.columns);
            thisObj.columnSorter();

            binningCont.insert('p', ':first-child').classed('dropdown binColumn none', true).attr('data-id', '-1').text('None');
            d3.select('.menuBinningby').on('click', function (): void {
                thisObj.hideMenus(thisObj.$binningCont);
            });

            d3.selectAll('.binColumn').on('click', function (): void {
                d3.selectAll('.binColumn').style('border', '0');
                thisObj.cacheOldMenuState();
                thisObj.binColumn = parseInt(d3.select(this).attr('data-id'), 10);
                thisObj.persistOrient();
                thisObj.setOrient(thisObj.options);

                thisObj.$binningCont.slideToggle();
            });
            $('.menuBinningby p').text(thisObj.viewModel.columns[thisObj.binColumn] ?
                thisObj.viewModel.columns[thisObj.binColumn].displayName : 'None')
                .attr('title', thisObj.viewModel.columns[thisObj.binColumn] ?
                    thisObj.viewModel.columns[thisObj.binColumn].displayName : 'None');
            if (thisObj.binColumn === -1) {
                d3.select('.binningCont .none').style('background-color', '#4c4c4c');
            } else {

                const names: string[] = thisObj.viewModel.columns[thisObj.binColumn]
                    .displayName.split(' ', thisObj.viewModel.columns.length);
                const jointName: string = names.join('.');
                d3.select(`.binningCont .${jointName}.index${thisObj.binColumn}`)
                    .style('background-color', '#4c4c4c');
            }
            /*For Y Axis Dropdown*/
            thisObj.renderYAxis(thisObj.topCont, 'groupingCont', 'menuY', 'groupedColumn');

            /*For X Axis Dropdown*/
            thisObj.renderXAxis(thisObj.topCont, 'targetCont', 'menuX', 'targetColumn');

            //color By
            const $menuColorBy: JQuery = $('.menuColorBy');
            // tslint:disable-next-line:no-any
            const colorByCont: any = thisObj.topCont.append('div').classed('presentationCont colorByCont', true)
                .style('left', $menuColorBy.position().left + px)
                .style('top', (($menuColorBy.position().top + thisObj.top) + px));
            thisObj.$colorByCont = $('.colorByCont');
            colorByCont.selectAll('p')
                .data(thisObj.viewModel.columns).enter()
                .append('p')
                .attr({
                    class: function (datum: DataViewMetadataColumn, index: number): string {
                        thisObj.previouscolorby = thisObj.colorByColumn;

                        return `dropdown colorByColumn ${datum.displayName} index${index}`;
                    },
                    'data-id': function (datum: DataViewMetadataColumn, index: number): number { return index; }
                })
                .text(function (datum: DataViewMetadataColumn): string { return datum.displayName; });

            //Sorting the columns in binning dropdown
            thisObj.dropdown = colorByCont.selectAll(`.dropdown`).data(thisObj.viewModel.columns);
            thisObj.columnSorter();

            colorByCont.insert('p', ':first-child').classed('dropdown colorByColumn none', true).attr('data-id', '-1').text('None');

            d3.select('.menuColorBy').on('click', function (): void {
                thisObj.hideMenus(thisObj.$colorByCont);
            });

            d3.selectAll('.colorByColumn').on('click', function (): void {
                d3.selectAll('.colorByColumn').style('border', '0');
                thisObj.cacheOldMenuState();
                thisObj.colorByColumn = parseInt(d3.select(this).attr('data-id'), 10);
                thisObj.persistOrient();
                thisObj.setOrient(thisObj.options);
                thisObj.$colorByCont.slideToggle();
            });

            $('.menuColorBy p').text(thisObj.viewModel.columns[thisObj.colorByColumn] ?
                thisObj.viewModel.columns[thisObj.colorByColumn].displayName : 'None')
                .attr('title', thisObj.viewModel.columns[thisObj.colorByColumn] ?
                    thisObj.viewModel.columns[thisObj.colorByColumn].displayName : 'None');
            if (thisObj.colorByColumn === -1) {
                d3.select('.colorByCont .none').style('background-color', '#4c4c4c');
            } else {

                const names: string[] = thisObj.viewModel.columns[thisObj.colorByColumn]
                    .displayName.split(' ', thisObj.viewModel.columns.length);
                const jointName: string = names.join('.');
                d3.select(`.colorByCont .${jointName}.index${thisObj.colorByColumn}`)
                    .style('background-color', '#4c4c4c');
            }
            /*For Y Axis*/
            thisObj.yAxisCont
                .append('p')
                .classed('presentationCont yAxisLabel', true)
                .style('max-width', `${$('.yAxisCont').height() - 10}px`)
                .append('p')
                .classed('yLabel', true)
                .text(`Y-axis: ${thisObj.viewModel.columns[thisObj.chartType.toLowerCase() === 'bar' ?
                    thisObj.groupedColumn : thisObj.targetColumn].displayName}`);

            d3.select('.presentationCont.yAxisLabel')
                .append('p')
                .classed('ySymbol', true)
                .text('▼');

            d3.select('.presentationCont.yAxisLabel')
                .style('top', function (): string {
                    // tslint:disable-next-line:no-any
                    const scroll: any = $('.presentationCont.yAxisLabel').height();

                    return `${(thisObj.height / 2) - (scroll / 2)}px`;
                })
                .style('left', '20px')
                .style('width', `${$('.yLabel').width() + $('.ySymbol').width() + 30}px`);

            if ($('.presentationCont.yAxisLabel').width() === ($('.yAxisCont').height() - 10)) {
                d3.select('.yLabel').style({
                    width: `${$('.yAxisCont').height() - 70}px`,
                    'text-overflow': 'ellipsis',
                    'white-space': 'nowrap',
                    overflow: 'hidden'
                });
                d3.select('.presentationCont.yAxisLabel').style('width', `${$('.yLabel').width() + $('.ySymbol').width() + 30}px`);
            }

            thisObj.$yAxisLabel = $('.yAxisLabel');
            thisObj.renderYAxis(thisObj.yAxisCont, 'yAxis', 'yAxisLabel', 'yAxisColumn');
            /*For X Axis*/
            thisObj.xAxisCont
                .append('p')
                .classed('presentationCont xAxisLabel', true)
                .append('p')
                .classed('xLabel', true)
                .text(`X-axis: ${thisObj.viewModel.columns[thisObj.chartType.toLowerCase() === 'bar' ?
                    thisObj.targetColumn : thisObj.groupedColumn].displayName}`);
            d3.select('.presentationCont.xAxisLabel')
                .style('margin-left', function (): string {
                    let parentWidth : number;
                    parentWidth = $('.presentationCont.xAxisLabel').width();

                    return `${(thisObj.width / 2) - (parentWidth / 2) - 40}px`;
                })
                .style('width', `${$('.xLabel').width() + $('.xSymbol').width() + 30}px`)
                .append('p')
                .classed('xSymbol', true)
                .text('▲');

            if ($('.presentationCont.xAxisLabel').width() === ($('.xAxisCont').width() - 40)) {
                d3.select('.xLabel').style({
                    width: `${$('.xAxisCont').width() - 70}px`,
                    'text-overflow': 'ellipsis',
                    'white-space': 'nowrap',
                    overflow: 'hidden'
                });
                d3.select('.presentationCont.xAxisLabel').style('width', `${$('.xLabel').width() + $('.xSymbol').width() + 30}px`);
            }

            thisObj.$xAxisLabel = $('.xAxisLabel');
            thisObj.renderXAxis(thisObj.xAxisCont, 'xAxis', 'xAxisLabel', 'xAxisColumn');

            d3.select('.menuColor').on('click', function (): void {
                thisObj.hideMenus(thisObj.$colorContShape);
            });

            d3.select('.menuTextColor').on('click', function (): void {
                thisObj.hideMenus(thisObj.$colorContText);
            });
            if (thisObj.selectionIndexes.length > 0) {
                thisObj.$xAxis.hide();
                thisObj.$yAxis.hide();
            }

            // Color Palette
            const $menuColor: JQuery = $('.menuColor');
            const $menuTextColor: JQuery = $('.menuTextColor');
            // tslint:disable-next-line: no-any
            const colorCont: any = thisObj.topCont.append('div').classed('colorCont shape', true)
                .style('left', $menuColor.position().left + px)
                .style('top', (($menuColor.position().top + thisObj.top) + px));
            // tslint:disable-next-line:no-any
            const colorCont2: any = thisObj.topCont.append('div').classed('colorCont text', true)
                .style('left', $menuTextColor.position().left + px)
                .style('top', (($menuTextColor.position().top + thisObj.top) + px));
            // tslint:disable-next-line:no-any
            const $colorCont: any = d3.selectAll('.colorCont');
            thisObj.$colorCont = $('.colorCont');
            thisObj.$colorContShape = $('.colorCont.shape');
            thisObj.$colorContText = $('.colorCont.text');
            const nospace: string = '';
            const cellName: string = 'cell';
            let iCounter: number;
            let jCounter: number;
            let kCounter: number;
            kCounter = 0;
            for (iCounter = 0; iCounter < 9; iCounter++) {
                $colorCont.append('tr');
                for (jCounter = 0; jCounter < 13; jCounter++) {
                    $colorCont.append('td').attr('class', cellName + iCounter + jCounter);
                    d3.selectAll(thisObj.dot + cellName + iCounter + jCounter)
                        .style('background-color', thisObj.colors[iCounter][jCounter])
                        .attr({
                            icounter: iCounter,
                            jcounter: jCounter
                        });
                    kCounter++;

                    colorCont.select(thisObj.dot + cellName + iCounter + jCounter)
                        .on('click', function (): void {
                            thisObj.cacheOldMenuState();
                            thisObj.iColumn = parseInt(d3.select(this).attr('icounter'), 10);
                            thisObj.jColumn = parseInt(d3.select(this).attr('jcounter'), 10);
                            thisObj.persistOrient();
                            thisObj.setOrient(thisObj.options);

                        })
                        .on('mouseover', function (): void {
                            d3.select(this)
                                // on hover line should be hightlighted hence 3
                                .style('border', '1px solid');
                        })
                        .on('mouseout', function (): void {
                            d3.select(this)
                                // on hover line should be hightlighted hence 3
                                .style('border', 'none');
                        });

                    colorCont2.select(thisObj.dot + cellName + iCounter + jCounter)
                        .on('click', function (): void {
                            thisObj.cacheOldMenuState();
                            thisObj.iColumntext = parseInt(d3.select(this).attr('icounter'), 10);
                            thisObj.jColumntext = parseInt(d3.select(this).attr('jcounter'), 10);
                            thisObj.persistOrient();
                            thisObj.setOrient(thisObj.options);

                        })
                        .on('mouseover', function (): void {
                            d3.select(this)
                                // on hover line should be hightlighted hence 3
                                .style('border', '1px solid');
                        })
                        .on('mouseout', function (): void {
                            d3.select(this)
                                // on hover line should be hightlighted hence 3
                                .style('border', 'none');
                        });
                }
            }
            //Highlighting selected colors
            colorCont.select(thisObj.dot + cellName + thisObj.iColumn + thisObj.jColumn)
                .style('border', '2px solid #000');

            colorCont2.select(thisObj.dot + cellName + thisObj.iColumntext + thisObj.jColumntext)
                .style('border', '2px solid #000');

            d3.selectAll('.menuColor p')
                .style('padding-top', '5px')
                .append('tr')
                .append('td')
                .style('height', '5px')
                .style({
                    'background-color': thisObj.colors[thisObj.iColumn][thisObj.jColumn]
                });

            d3.selectAll('.menuTextColor p')
                .style('padding-top', '5px')
                .append('tr')
                .append('td')
                .style('height', '5px')
                .style({
                    'background-color': thisObj.colors[thisObj.iColumntext][thisObj.jColumntext]
                });
            // Dynamic Binning UI part
            //here
            d3.select('.menuRangeforBinning')
                .on('click', function (): void {
                    thisObj.hideMenus(thisObj.$dynamicBinCont);
                });

            const $menuRangeforBinning: JQuery = $('.menuRangeforBinning');
            // tslint:disable-next-line:no-any
            const dynamicBinCont: any = thisObj.topCont.append('div').classed('dynamicBinCont', true)
                .style('left', $menuRangeforBinning.position().left + px)
                .style('top', (($menuRangeforBinning.position().top + thisObj.top) + px));
            thisObj.$dynamicBinCont = $('.dynamicBinCont');
            const colData: string[] = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'];
            dynamicBinCont.selectAll('p')
                .data(colData).enter()
                .append('p')
                .attr({
                    class: function (datum: string, index: number): string {
                        return `col${index + 1}`;
                    },
                    'data-id': function (datum: string, index: number): number { return index + 1; }
                })
                .text(function (datum: string, index: number): number { return (index + 1); })
                .classed('dynamicBinColumn', true);
            d3.selectAll('.dynamicBinColumn').on('click', function (): void {
                d3.selectAll('.dynamicBinColumn').style('border', '0');
                thisObj.numberOfBins = parseInt(d3.select(this).attr('data-id'), 10);
                thisObj.persistOrient();
                thisObj.setOrient(thisObj.options);
                $('.menuRangeforBinning p').text(thisObj.numberOfBins);
                $('.dynamicBinCont').slideToggle();
                d3.select(`.dynamicBinCont .col${thisObj.numberOfBins}`).style('border', '1px solid');
            });
            $('.menuRangeforBinning p').text(thisObj.numberOfBins);
            d3.select(`.dynamicBinCont .col${thisObj.numberOfBins}`).style('background-color', '#4c4c4c');
        }
        private cacheOldMenuState(): void {
            const thisObj: this = this;
            thisObj.groupedColumnOld.push(thisObj.groupedColumn);
            thisObj.targetColumnOld.push(thisObj.targetColumn);
            thisObj.binColumnOld.push(thisObj.binColumn);
            thisObj.colorByColumnOld.push(thisObj.colorByColumn);
            thisObj.chartTypeOld.push(thisObj.chartType);
            thisObj.iColumnold.push(thisObj.iColumn);
            thisObj.jColumnOld.push(thisObj.jColumn);
            thisObj.iColumntextold.push(thisObj.iColumntext);
            thisObj.jColumntextOld.push(thisObj.jColumntext);
            if (!thisObj.type) {
                thisObj.actions.push(1);
            }
        }

        private cacheSelectionState(): void {
            const thisObj: this = this;
            if (0 === thisObj.selectionIndexes.length && 0 === thisObj.globalSelections.length) {
                return;
            }
            thisObj.selectionsOld.push(thisObj.selectionIndexes.slice(0));
            thisObj.globalSelectionsOld.push(thisObj.globalSelections.slice(0));
            if (!thisObj.type) {
                thisObj.actions.push(2);
            }
        }
        // tslint:disable-next-line:no-any
        private renderBrickChart(data: any, index: number, categories: any,
                                 prevLength: number, chartWidth: number, height: number, marginLeft: number): void {
            let thisObj: this;
            thisObj = this;
            const brick: string = 'brick';
            thisObj.$xAxisLabel.show();
            thisObj.$yAxisLabel.show();
            // tslint:disable-next-line:no-any
            let brickChart1: any;
            // tslint:disable-next-line:no-any
            let brickChart: any;
            if (thisObj.binColumn === -1) {
                brickChart1 = thisObj.mainCont
                    .append('div')
                    .style('margin', '30px')
                    .classed('brickChart1', true)

                    .style({
                        'text-align': 'center'
                    })
                    .style('margin-left', marginLeft + thisObj.px);

            } else {
                brickChart1 = thisObj.mainCont
                    .append('div')
                    .classed('brickChart1', true)
                    .classed('borderChart', true)
                    .style({
                        'text-align': 'center'
                    });

            }
            brickChart1.append('label')
                .text(thisObj.returnLabelText(categories[index]))
                .style('font-weight', 'bold');

            brickChart = brickChart1
                .append('div')
                .attr('class', 'brickChart')
                .style('height', ('120px'))
                .style('margin-left', marginLeft + thisObj.px)
                .style({
                    margin: '0px',
                    border: 'none',
                    overflow: 'auto'
                })
                .style('margin-top', function (): string {

                    if (categories[index] !== '') {
                        return '0px';
                    } else {
                        return '15px';
                    }

                })
                //width of chart is reduced by 10px to reduce extra width which was causing the x-axis scroll problem
                .style('width', $('.brickChart1').width() - 10 + thisObj.px);
            let percentSign: string;
            percentSign = '%';
            let brickWidth: number;
            brickWidth = $('.brickChart').width();
            let sum: number;
            let increment: number;
            let value: number;
            increment = thisObj.previousDataLength;
            brickChart.selectAll('div')
                .data(data)
                .enter().append('div')
                .classed(`brick index${index}`, true)
                .style({
                    margin: '1px',
                    padding: '0'
                })
                .style('background-color', function (): string {
                    if (value === undefined || sum === 0 || value === 0) {
                        return thisObj.colors[thisObj.iColumn][thisObj.jColumn];
                    } else {
                        return 'transparent';
                    }
                })
                // tslint:disable-next-line:no-any
                .each(function (datum: any): void {
                    thisObj.globalSelections.push({
                        data: datum,
                        data1: datum.values[0],
                        binIndex: index,
                        category: categories[index]
                    });
                    if (thisObj.colorByColumn !== -1 && !thisObj.isCategory) {
                        // tslint:disable-next-line:no-any
                        const THIS: any = this;
                        sum = thisObj.getSum(datum);
                        thisObj.globalSelections.push({
                            data: datum,
                            data1: datum.values[0],
                            binIndex: index,
                            category: categories[index]
                        });
                        // tslint:disable-next-line:no-any
                        datum.values[0].forEach(function (d: any, i: number): any {
                            value = d[`values`][`value`];
                            if (value > 0) {
                                thisObj.isCategory = false;
                            }
                            let percent: number;
                            percent = d[`values`][`value`] / sum;
                            const p: number = d[`values`][`value`] / sum;
                            if (thisObj.colorByColumn !== thisObj.binColumn) {
                                d3.select(THIS)
                                    .append('div')
                                    .classed('inDiv', true)
                                    .style('float', 'left')
                                    .style('margin', '0')
                                    .style('height', '100%')
                                    .style('padding', '0')
                                    .style('width', (percent * 55) + thisObj.px)
                                    .each(function (): void {
                                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                                thisObj.getTooltipData(tooltipEvent.data, d, 'brick', '', 0, thisObj),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);
                                    })
                                    // tslint:disable-next-line:no-any
                                    .attr('data-selection', function (dat: any, iterator: number): number { return increment; })
                                    // tslint:disable-next-line:no-any
                                    .style('background-color', function (): any {
                                        if (sum === 0) {
                                            return thisObj.colors[thisObj.iColumn][thisObj.jColumn];
                                        } else {
                                            // tslint:disable-next-line:no-any
                                            return thisObj.barcolor.filter(function (v: any): any {
                                                return v.key.toString() === d[`key`].toString();
                                            })[0].value;

                                        }
                                    });
                                increment++;
                            } else {
                                d3.select(THIS)
                                    .append('div')
                                    .classed('inDiv', true)
                                    .style('float', 'left')
                                    .style('margin', '0')
                                    .style('height', '100%')
                                    .style('padding', '0')
                                    .style('width', percent * 55 + thisObj.px)
                                    .each(function (): void {
                                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                                thisObj.getTooltipData(tooltipEvent.data, d, 'brick', '', 0, thisObj),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);
                                    })
                                    // tslint:disable-next-line:no-any
                                    .attr('data-selection', function (dat: any, iterator: number): number { return increment; })
                                    // tslint:disable-next-line:no-any
                                    .style('background-color', function (): any {
                                        if (sum === 0) {
                                            return thisObj.colors[thisObj.iColumn][thisObj.jColumn];
                                        } else {
                                            // tslint:disable-next-line:no-any
                                            return thisObj.barcolor.filter(function (v: any): any {
                                                return v.key.toString() === d[`key`].toString();
                                            })[0].value;
                                        }
                                    });
                                increment++;
                            }
                        });
                    } else {
                        brickChart.selectAll('div')
                            .data(data)
                            .enter().append('div')
                            .classed(`brick index${index}`, true);
                        d3.selectAll('div')
                            // tslint:disable-next-line:no-any
                            .each(function (tdatum: any): void {
                                thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                         (tooltipEvent: TooltipEventArgs<number>) =>
                                        thisObj.getTooltipData(tooltipEvent.data, tdatum, 'bar', '', 0, thisObj),
                                                                         (tooltipEvent: TooltipEventArgs<number>) => null);
                            });
                    }
                })
                // tslint:disable-next-line:no-any
                .attr('data-selection', function (datum: any, iterator: number): number { return iterator + increment; })
                .style({
                    height: `${thisObj.textSize * 2 - 2}px`,
                    width: `${Math.floor(thisObj.textWidth + 14)}px`
                })
                // tslint:disable-next-line:no-any
                .style('background-color', function (): any {
                    if (value === undefined || value === 0) {
                        return thisObj.colors[thisObj.iColumn][thisObj.jColumn];
                    } else {
                        return 'transparent';
                    }
                })

                .append('p')
                // tslint:disable-next-line:no-any
                .attr('title', function (datum: any): any {
                    return datum[`key`].toString() || 'Blank';
                })

                .style({
                    'font-size': `${this.settings.fontSettings.fontSize}px`,
                    'font-family': this.settings.fontSettings.fontFamily,
                    'white-space': 'nowrap',
                    overflow: 'hidden',
                    height: `${thisObj.textSize}px`
                })
                // tslint:disable-next-line:no-any
                .text(function (tdatum: any): any {
                    let labelText: string;
                    labelText = tdatum[`key`].toString() || 'Blank';

                    return labelText;
                })

                .style('transform', 'scale(0)')
                .transition()
                .duration(2000)
                .style('transform', 'scale(1)');

            d3.selectAll(`.brick.index${index}`).append('p')
                .style('text-align', 'right')
                .style('top', `${thisObj.textSize}px`)
                .style('font-size', '10px')
                .style('width', '35px')
                .style('margin-left', '20px')
                .classed('brickVal', true)
                .style({
                    'font-size': `${this.settings.fontSettings.fontSize}px`,
                    'font-family': this.settings.fontSettings.fontFamily
                })

                // tslint:disable-next-line:no-any
                .text(function (datum: any): string {
                    if (thisObj.colorByColumn !== -1) {
                        sum = thisObj.getSum(datum);
                    } else {
                        sum = datum[`values`][`value`];
                    }

                    // tslint:disable-next-line:no-any
                    let displayVal: any;
                    let tempMeasureData: string = sum.toString();
                    const valLen: number = tempMeasureData.split('.')[0].length;
                    displayVal = thisObj.returnDisplayVal(valLen);

                    //Value Formatter Creation
                    const formatter: IValueFormatter = ValueFormatter.create({
                        format: thisObj.dataView.metadata.columns[thisObj.targetColumn].format,
                        value: displayVal
                    });
                    tempMeasureData = formatter.format(sum);

                    return tempMeasureData;
                })
                //for  p label
                .style('background-color', 'transparent')
                .style('color', thisObj.colors[thisObj.iColumntext][thisObj.jColumntext]);

            d3.selectAll('.brickChart div').style('transform', 'scale(0)').transition().duration(500).style('transform', 'scale(1)');
            brickChart.style('transform', 'scale(0)')
                .transition()
                .duration(1000)
                .style('transform', 'scale(1)');
        }
        // tslint:disable-next-line:no-any
        private renderBinBrickChart(data: any, index: number, categories: any,
                                    chartWidth: number, height: number, marginLeft: number): void {
            let thisObj: this;
            thisObj = this;
            const brick: string = 'brick';
            thisObj.$xAxisLabel.show();
            thisObj.$yAxisLabel.show();

            // tslint:disable-next-line:no-any
            let brickChart1: any;
            // tslint:disable-next-line:no-any
            let brickChart: any;
            let value: number;
            brickChart1 = thisObj.mainCont
                .append('div')
                .classed('brickChart1', true)
                .classed('borderChart', true)
                .style({
                    overflow: 'hidden',
                    'text-align': 'center'
                })
                .style('margin-left', marginLeft + thisObj.px);

            brickChart1.append('label')
                .text(thisObj.returnLabelText(categories[index]))
                .style('font-weight', 'bold');

            brickChart = brickChart1
                .append('div')
                .attr('class', 'brickChart')
                .style('height', ('120px'))
                .style('margin-left', marginLeft + thisObj.px)
                .style({
                    margin: '0px',
                    border: 'none',
                    overflow: 'auto'
                })
                .style('margin-top', function (): string {

                    if (categories[index] !== '') {
                        return '0px';
                    } else {
                        return '15px';
                    }

                })
                //width of chart is reduced by 10px to reduce extra width which was causing the x-axis scroll problem
                .style('width', $('.brickChart1').width() - 10 + thisObj.px);
            let percentSign: string;
            percentSign = '%';
            const $brickChart: JQuery = $('.brickChart');
            const brickChartWidth: number = $brickChart.width();
            let sum: number;

            let increment: number;
            increment = thisObj.previousDataLength;
            brickChart.selectAll('div')
                .data(data)
                .enter().append('div')
                .classed(`brick index${index}`, true)
                .style({
                    margin: '1px',
                    padding: '0'
                })
                .style('background-color', function (): string {
                    if (value === undefined || sum === 0 || value === 0) {
                        return thisObj.colors[thisObj.iColumn][thisObj.jColumn];
                    } else {
                        return 'transparent';
                    }
                })
                // tslint:disable-next-line:no-any
                .each(function (datum: any): void {
                    thisObj.globalSelections.push({
                        data: datum,
                        data1: datum.values[0],
                        binIndex: index,
                        category: categories[index]
                    });
                    if (thisObj.colorByColumn !== -1 && !thisObj.isCategory) {
                        // tslint:disable-next-line:no-any
                        const THIS: any = this;
                        sum = thisObj.getSum(datum);
                        thisObj.globalSelections.push({
                            data: datum,
                            data1: datum.values[0],
                            binIndex: index,
                            category: categories[index]
                        });
                        // tslint:disable-next-line:no-any
                        datum.values[0].forEach(function (d: any, i: number): any {
                            value = d[`values`][`value`];
                            if (value > 0) {
                                thisObj.isCategory = false;
                            }
                            thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                     (tooltipEvent: TooltipEventArgs<number>) =>
                                    thisObj.getTooltipData(tooltipEvent.data, d, 'brick', '', 0, thisObj),
                                                                     (tooltipEvent: TooltipEventArgs<number>) => null);
                            let percent: number;
                            percent = d[`values`][`value`] / sum;
                            const p: number = d[`values`][`value`] / sum;
                            if (thisObj.colorByColumn !== thisObj.binColumn) {
                                d3.select(THIS)
                                    .append('div')
                                    .classed('inDiv', true)
                                    .style('float', 'left')
                                    .style('margin', '0')
                                    .style('height', '100%')
                                    .style('padding', '0')
                                    .style('width', percent * 55 + thisObj.px)
                                    .each(function (): void {
                                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                                thisObj.getTooltipData(tooltipEvent.data, d, 'brick', '', 0, thisObj),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);
                                    })
                                    // tslint:disable-next-line:no-any
                                    .attr('data-selection', function (dat: any, iterator: number): number { return increment; })
                                    // tslint:disable-next-line:no-any
                                    .style('background-color', function (): any {

                                        if (value === 0) {
                                            return thisObj.colors[thisObj.iColumn][thisObj.jColumn];
                                        } else {

                                            // tslint:disable-next-line:no-any
                                            return thisObj.barcolor.filter(function (v: any): any {
                                                return v.key.toString() === d[`key`].toString();
                                            })[0].value;
                                        }
                                    });
                                increment++;
                            } else {
                                d3.select(THIS)
                                    .append('div')
                                    .classed('inDiv', true)
                                    .style('float', 'left')
                                    .style('margin', '0')
                                    .style('height', '100%')
                                    .style('padding', '0')
                                    .style('width', percent * 55 + thisObj.px)
                                    .each(function (): void {
                                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                                thisObj.getTooltipData(tooltipEvent.data, d, 'brick', '', 0, thisObj),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);
                                    })
                                    // tslint:disable-next-line:no-any
                                    .attr('data-selection', function (dat: any, iterator: number): number { return increment; })
                                    // tslint:disable-next-line:no-any
                                    .style('background-color', function (): any {

                                        if (value === 0) {
                                            return thisObj.colors[thisObj.iColumn][thisObj.jColumn];
                                        } else {

                                            // tslint:disable-next-line:no-any
                                            return thisObj.barcolor.filter(function (v: any): any {
                                                return v.key.toString() === d[`key`].toString();
                                            })[0].value;
                                        }
                                    });
                                increment++;
                            }
                        });
                    } else {
                        brickChart.selectAll('div')
                            .data(data)
                            .enter().append('div')
                            .classed(`brick index${index}`, true);
                        d3.selectAll('div')
                            // tslint:disable-next-line:no-any
                            .each(function (tdatum: any): void {
                                thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                         (tooltipEvent: TooltipEventArgs<number>) =>
                                        thisObj.getTooltipData(tooltipEvent.data, tdatum, 'bar', '', 0, thisObj),
                                                                         (tooltipEvent: TooltipEventArgs<number>) => null);
                            });
                    }

                })
                // tslint:disable-next-line:no-any
                .attr('data-selection', function (datum: any, iterator: number): number { return iterator + increment; })
                .style({
                    height: `${thisObj.textSize * 2 - 2}px`,
                    width: `${thisObj.textWidth + 14}px`
                })
                .style('background-color', function (): string {
                    if (value === undefined || value === 0) {
                        return thisObj.colors[thisObj.iColumn][thisObj.jColumn];
                    } else {
                        return 'transparent';
                    }
                })

                .append('p')
                // tslint:disable-next-line:no-any
                .attr('title', function (datum: any): any {
                    return datum[`key`].toString() || 'Blank';
                })
                // tslint:disable-next-line:no-any
                .style({
                    'font-size': `${this.settings.fontSettings.fontSize}px`,
                    'font-family': this.settings.fontSettings.fontFamily,
                    'white-space': 'nowrap',
                    overflow: 'hidden',

                    height: `${thisObj.textSize}px`
                })
                // tslint:disable-next-line:no-any
                .text(function (datum: any): any {
                    // tslint:disable-next-line:no-any
                    let labelText: any;
                    labelText = datum[`key`].toString() || 'Blank';

                    return labelText;
                })
                .style('transform', 'scale(0)')
                .transition()
                .duration(2000)
                .style('transform', 'scale(1)');

            d3.selectAll(`.brick.index${index}`).append('p')
                .style('text-align', 'right')
                .style('top', '12px')
                .style('font-size', '10px')
                .style('width', '35px')
                .style('margin-left', '20px')
                .classed('brickVal', true)
                .style({
                    'font-size': `${this.settings.fontSettings.fontSize}px`,
                    'font-family': this.settings.fontSettings.fontFamily
                })

                // tslint:disable-next-line:no-any
                .text(function (datum: any): string {
                    if (thisObj.colorByColumn !== -1) {
                        sum = thisObj.getSum(datum);
                    } else {
                        sum = datum[`values`][`value`];
                    }

                    // tslint:disable-next-line:no-any
                    let displayVal: any;
                    let tempMeasureData: string = sum.toString();
                    const valLen: number = tempMeasureData.split('.')[0].length;
                    displayVal = thisObj.returnDisplayVal(valLen);

                    //Value Formatter Creation
                    const formatter: IValueFormatter = ValueFormatter.create({
                        format: thisObj.dataView.metadata.columns[thisObj.targetColumn].format,
                        value: displayVal
                    });
                    tempMeasureData = formatter.format(sum);

                    return tempMeasureData;
                })
                .style('background-color', 'transparent')
                .style('color', thisObj.colors[thisObj.iColumntext][thisObj.jColumntext]);

            d3.selectAll('.brickChart div').style('transform', 'scale(0)').transition().duration(500).style('transform', 'scale(1)');
            brickChart.style('transform', 'scale(0)')
                .transition()
                .duration(1000)
                .style('transform', 'scale(1)');
            // thisObj.createSelectionBox(thisObj.mainCont, '.mainCont', '.brickChart > div');
        }
        // tslint:disable-next-line:no-any
        private renderBinBarChart(data: any, index: number,
            // tslint:disable-next-line:no-any
                                  categories: any, chartWidth: number, chartHeight: number, height: number, marginLeft: number): void {
            this.noOfBars = data.length;
            let thisObj: this;
            thisObj = this;
            thisObj.isCategory = true;
            // tslint:disable-next-line:no-any
            let chart: any;
            // tslint:disable-next-line:no-any
            let chart1: any;
            // tslint:disable-next-line:no-any
            let subDiv1: any;
            const indexString: string = 'index';
            const len: number = data.length;
            const categoriesLength: number = categories.length;
            // tslint:disable-next-line:no-any
            let subDiv: any;
            const labelHeight: number = 16;
            const labelMarginBottom: number = 10;
            const marginBottom: number = 10;
            const marginBetweenBars: number = 10;
            const percentConverter: number = 100;
            // Ratio
            const ratio: number = thisObj.settings.ratio.percent / percentConverter;
            let rowWidth: number;

            if (thisObj.colorByColumn !== -1) {
                thisObj.getMaxValue(data);
            } else {
                thisObj.getMaxValue1(data);
            }
            const margin: number = 30;
            chart = thisObj.mainCont
                .append('div')
                .attr('class', 'chart borderChart binChart')
                .style({
                    position: 'relative',
                    width: chartWidth + thisObj.px,
                    height: chartHeight + thisObj.px,
                    'margin-left': marginLeft + thisObj.px,
                    overflow: 'hidden'
                });
            subDiv1 = chart.append('div')
                .style({
                    'min-height': '10px',
                    width: '100%',
                    padding: '0px',
                    margin: '0px',
                    'text-align': 'center',
                    'font-weight': 'bold',
                    height: '20px'
                });

            subDiv1.append('label')
                .text(thisObj.returnLabelText(categories[index]) || 'Blank')
                .attr('title', thisObj.returnLabelText(categories[index]))
                .classed('label', true);

            chart1 = chart.append('div')
                .style({
                    position: 'relative',
                    'max-height': '100%',
                    overflow: 'auto'
                });

            subDiv = chart1.append('div')
                .classed('subDiv', true)
                .classed('innerdiv', true)
                .style({
                    width: chartWidth + thisObj.px,
                    'min-height': chartHeight + thisObj.px,
                    padding: '0',
                    margin: '0',
                    overflow: 'auto'

                });

            subDiv.append('label');

            let increment: number;
            increment = thisObj.previousDataLength;
            const percentSign: string = '%';
            const row: string = 'row';

            // tslint:disable-next-line:no-any
            let tooltipchart: any;
            tooltipchart = subDiv.selectAll('div')
                .data(data)
                .enter()
                .append('div')
                .classed(row + index, true)
                .classed('row', true);

            rowWidth = parseInt(d3.select('.row').style('width').split('p')[0], 10) +
                parseInt(d3.select('.row').style('margin-left').split('p')[0], 10);

            d3.selectAll(thisObj.dot + row + index)
                .append('p')
                .classed('categoryText', true)
                // tslint:disable-next-line:no-any
                .attr('title', function (datum: any): string { return datum[`key`].toString(); })
                // tslint:disable-next-line:no-any
                .style('margin-bottom', function (datum: any, iterator: number): string {
                    if (iterator === data.length - 1) {
                        return marginBottom + thisObj.px;
                    }

                    return '1px';
                })
                // tslint:disable-next-line:no-any
                .text(function (datum: any): string {

                    let labelText: string;
                    labelText = datum[`key`].toString() || 'Blank';

                    return labelText;
                });
            let sum: number = 0;
            // tslint:disable-next-line:no-any
            const tooltipData: any = [];
            d3.selectAll(thisObj.dot + row + index).append('div')
                .classed('thebar', true)

                // tslint:disable-next-line:no-any
                .attr('data-selection', function (datum: any, iterator: number): number { return iterator + increment; })
                // tslint:disable-next-line:no-any
                .each(function (datum: any): any {
                    if (thisObj.colorByColumn !== -1) {
                        // tslint:disable-next-line:no-any
                        const THIS: any = this;
                        sum = thisObj.getSum(datum);
                        // tslint:disable-next-line:no-any
                        datum.values[0].forEach(function (d: any, i: number): any {
                            thisObj.globalSelections.push({
                                data: datum,
                                data1: datum.values[0],
                                binIndex: index,
                                category: categories[index]
                            });

                            const value: number = d[`values`][`value`];
                            if (value > 0) {
                                thisObj.isCategory = false;
                            }
                            thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                     (tooltipEvent: TooltipEventArgs<number>) =>
                                    thisObj.getTooltipData(tooltipEvent.data, d, 'bar', '', 0, thisObj),
                                                                     (tooltipEvent: TooltipEventArgs<number>) => null);
                            const p: number = d[`values`][`value`] / sum;

                            tooltipData.push({ key: d[`key`], value: d[`values`][`value`] });
                            if (thisObj.colorByColumn !== thisObj.binColumn) {
                                d3.select(THIS)
                                    .append('div').classed('inDiv', true)
                                    .style('float', 'left')
                                    .style('margin', '0')
                                    .style('width', d[`values`][`value`] <= 0 ? `0` :
                                        // tslint:disable-next-line:max-line-length
                                        `${p * ((sum / thisObj.maxValue) * (thisObj.binColumn === -1 ? 0.85 : 0.60) * (1 - ratio) * rowWidth)}px`)
                                    // tslint:disable-next-line:no-any
                                    .on('click', function (dat: any): void {
                                        const selectionlen: number = dat.values[0].length;
                                        let level: number = 0;
                                        for (i = 0; i < selectionlen; i++) {
                                            if (dat.values[0][i].key === d.key) {
                                                level = i;
                                                break;
                                            }
                                        }
                                        // tslint:disable-next-line:no-any
                                        const selectionIds: any[] = dat.values[0][level].values.selectionId;
                                        // tslint:disable-next-line:no-any
                                        const lastindex: any = thisObj.selectionManager;
                                        // tslint:disable-next-line:no-any
                                        const lastIndexLen: any = lastindex.selectedIds.length;
                                        const isCrtlPressed: boolean = (d3.event as MouseEvent).ctrlKey;
                                        thisObj.selectionManager.select(selectionIds, isCrtlPressed).then((ids: ISelectionId[]) => {
                                            if (isCrtlPressed) {
                                            if (lastIndexLen === 0 && !(d3.select(this).classed('selected'))) {
                                                d3.selectAll('.inDiv').style({
                                                    opacity: 0.3
                                                });
                                                d3.select(this).style({
                                                    opacity: 1
                                                })
                                                .classed('selected', true);
                                                thisObj.lastSelectedbar = 1;
                                            } else if (d3.select(this).classed('selected')) {
                                                d3.select(this).style({
                                                    opacity: 0.3
                                                })
                                                .classed('selected', false);
                                                thisObj.lastSelectedbar--;
                                                if (thisObj.lastSelectedbar === 0) {
                                                    d3.selectAll('.inDiv').style({
                                                        opacity: 1
                                                    });
                                                }
                                            } else {
                                                d3.select(this).style({
                                                    opacity: 1
                                                })
                                                .classed('selected', true);
                                                thisObj.lastSelectedbar++;
                                            }
                                        } else {
                                            if (d3.select(this).classed('selected')) {
                                                d3.select(this).classed('selected', false);
                                                thisObj.selectionManager.clear().then(() => d3.selectAll('.inDiv').style({
                                                    opacity: 1
                                                }));
                                                if (!isCrtlPressed) {
                                                    thisObj.lastSelectedbar--;
                                                }
                                            } else {
                                            d3.selectAll('.inDiv').style({
                                                opacity: 0.3
                                            })
                                            .classed('selected', false);
                                            d3.select(this).style({
                                                opacity: 1
                                            })
                                            .classed('selected', true);
                                            thisObj.lastSelectedbar = 0;
                                            if (!isCrtlPressed) {
                                                    thisObj.lastSelectedbar++;
                                                }
                                        }
                                        }
                                        });
                                        (<Event>d3.event).stopPropagation();
                                    })
                                    // tslint:disable-next-line:no-any
                                    .each(function (tdatum: any): void {
                                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                                thisObj.getTooltipData(tooltipEvent.data, d, 'bar', '', 0, thisObj),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);
                                    })
                                    // tslint:disable-next-line:no-any
                                    .attr('data-selection', function (dat: any, iterator: number): number { return increment; })

                                    // tslint:disable-next-line:no-any
                                    .style('background-color', thisObj.barcolor.filter(function (v: any): any {
                                        return v.key.toString() === d[`key`].toString();
                                    })[0].value);
                                // tslint:disable-next-line:typedef
                                subDiv.on('click', function () {
                                    thisObj.selectionManager.clear().then(() => d3.selectAll('.inDiv').style({
                                        opacity: 1
                                    }));
                                    d3.selectAll('.inDiv').classed('selected', false);
                                });
                                increment++;
                            } else {
                                d3.select(THIS)
                                    .append('div').classed('inDiv', true)
                                    .style('float', 'left')
                                    .style('margin', '0')
                                    .style('width', d[`values`][`value`] <= 0 ? `0` :
                                        // tslint:disable-next-line:max-line-length
                                        `${p * ((sum / thisObj.maxValue) * (thisObj.binColumn === -1 ? 0.85 : 0.60) * (1 - ratio) * rowWidth)}px`)
                                    // tslint:disable-next-line:no-any
                                    .on('click', function (dat: any): void {
                                        const selectionlen: number = datum.values[0].length;
                                        let level: number = 0;
                                        for (i = 0; i < selectionlen; i++) {
                                            if (datum.values[0][i].key === d.key) {
                                                level = i;
                                                break;
                                            }
                                        }
                                        // tslint:disable-next-line:no-any
                                        const selectionIds: any[] = datum.values[0][level].values.selectionId;
                                        // tslint:disable-next-line:no-any
                                        const lastindex: any = thisObj.selectionManager;
                                        // tslint:disable-next-line:no-any
                                        const lastIndexLen: any = lastindex.selectedIds.length;
                                        const isCrtlPressed: boolean = (d3.event as MouseEvent).ctrlKey;
                                        thisObj.selectionManager.select(selectionIds, isCrtlPressed).then((ids: ISelectionId[]) => {
                                            if (isCrtlPressed) {
                                                if (lastIndexLen === 0 && !(d3.select(this).classed('selected'))) {
                                                    d3.selectAll('.inDiv').style({
                                                        opacity: 0.3
                                                    });
                                                    d3.select(this).style({
                                                        opacity: 1
                                                    })
                                                    .classed('selected', true);
                                                    thisObj.lastSelectedbar = 1;
                                                } else if (d3.select(this).classed('selected')) {
                                                    d3.select(this).style({
                                                        opacity: 0.3
                                                    })
                                                    .classed('selected', false);
                                                    thisObj.lastSelectedbar--;
                                                    if (thisObj.lastSelectedbar === 0) {
                                                        d3.selectAll('.inDiv').style({
                                                            opacity: 1
                                                        });
                                                    }
                                                } else {
                                                    d3.select(this).style({
                                                        opacity: 1
                                                    })
                                                    .classed('selected', true);
                                                    thisObj.lastSelectedbar++;
                                                }
                                            } else {
                                                if (d3.select(this).classed('selected')) {
                                                    d3.select(this).classed('selected', false);
                                                    thisObj.selectionManager.clear().then(() => d3.selectAll('.inDiv').style({
                                                        opacity: 1
                                                    }));
                                                    if (!isCrtlPressed) {
                                                        thisObj.lastSelectedbar--;
                                                    }
                                                } else {
                                                d3.selectAll('.inDiv').style({
                                                    opacity: 0.3
                                                })
                                                .classed('selected', false);
                                                d3.select(this).style({
                                                    opacity: 1
                                                })
                                                .classed('selected', true);
                                                thisObj.lastSelectedbar = 0;
                                                if (!isCrtlPressed) {
                                                    thisObj.lastSelectedbar++;
                                                }
                                            }
                                            }
                                        });
                                        (<Event>d3.event).stopPropagation();
                                    })
                                    // tslint:disable-next-line:no-any
                                    .each(function (tdatum: any): void {
                                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                                thisObj.getTooltipData(tooltipEvent.data, tdatum, 'bar', '', 0, thisObj),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);
                                    })
                                    // tslint:disable-next-line:no-any
                                    .attr('data-selection', function (dat: any, iterator: number): number { return increment; })
                                    // tslint:disable-next-line:no-any
                                    .style('background-color', function (): any {
                                        if (value === 0) {
                                            return thisObj.colors[thisObj.iColumntext][thisObj.jColumntext];
                                        } else {
                                            // tslint:disable-next-line:no-any
                                            return thisObj.barcolor.filter(function (v: any): any {
                                                return v.key.toString() === d[`key`].toString();
                                            })[0].value;

                                        }
                                    });
                                // tslint:disable-next-line:typedef
                                subDiv.on('click', function () {
                                    thisObj.selectionManager.clear().then(() => d3.selectAll('.inDiv').style({
                                        opacity: 1
                                    }));
                                    d3.selectAll('.inDiv').classed('selected', false);
                                });
                                increment++;
                            }
                        });
                    } else {
                        // tslint:disable-next-line:no-any
                        const value: any = datum[`values`][`value`];
                        if (value > 0) {
                            thisObj.isCategory = false;
                        }
                        d3.select(this)
                            .style('width', datum[`values`][`value`] <= 0 ? `0` :
                                `${(parseInt(datum[`values`][`value`], 10) / thisObj.maxValue * (
                                    (thisObj.binColumn === -1 ? 0.85 : 0.60) * (1 - ratio) * rowWidth))}px`)
                            .style('background-color', thisObj.colors[thisObj.iColumn][thisObj.jColumn])
                            // tslint:disable-next-line:no-any
                            .on('click', function (d: any): void {
                                // tslint:disable-next-line:no-any
                                const selectionIds: any[] = datum[`values`][`selectionId`];
                                // tslint:disable-next-line:no-any
                                const lastindex: any = thisObj.selectionManager;
                                // tslint:disable-next-line:no-any
                                const lastIndexLen: any = lastindex.selectedIds.length;
                                const isCrtlPressed: boolean = (d3.event as MouseEvent).ctrlKey;
                                thisObj.selectionManager.select(selectionIds, isCrtlPressed).then((ids: ISelectionId[]) => {
                                    if (isCrtlPressed) {
                                    if (lastIndexLen === 0 && !(d3.select(this).classed('selected'))) {
                                        d3.selectAll('.thebar').style({
                                            opacity: 0.3
                                        });
                                        d3.select(this).style({
                                            opacity: 1
                                        })
                                        .classed('selected', true);
                                        thisObj.lastSelectedbar = 1;
                                    } else if (d3.select(this).classed('selected')) {
                                        d3.select(this).style({
                                            opacity: 0.3
                                        })
                                        .classed('selected', false);
                                        thisObj.lastSelectedbar--;
                                        if (thisObj.lastSelectedbar === 0) {
                                            d3.selectAll('.thebar').style({
                                                opacity: 1
                                            });
                                        }
                                    } else {
                                        d3.select(this).style({
                                            opacity: 1
                                        })
                                        .classed('selected', true);
                                        thisObj.lastSelectedbar++;
                                    }
                                } else {
                                    if (d3.select(this).classed('selected')) {
                                        d3.select(this).classed('selected', false);
                                        thisObj.selectionManager.clear().then(() => d3.selectAll('.thebar').style({
                                            opacity: 1
                                        }));
                                        if (!isCrtlPressed) {
                                            thisObj.lastSelectedbar--;
                                        }
                                    } else {
                                    d3.selectAll('.thebar').style({
                                        opacity: 0.3
                                    })
                                    .classed('selected', false);
                                    d3.select(this).style({
                                        opacity: 1
                                    })
                                    .classed('selected', true);
                                    thisObj.lastSelectedbar = 0;
                                    if (!isCrtlPressed) {
                                        thisObj.lastSelectedbar++;
                                    }
                                }
                                }
                                });
                                (<Event>d3.event).stopPropagation();
                            });

                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                thisObj.getTooltipData(tooltipEvent.data, datum, 'bin', '', 0, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);
                        // tslint:disable-next-line:typedef
                        subDiv.on('click', function () {
                            thisObj.selectionManager.clear().then(() => d3.selectAll('.thebar').style({
                                opacity: 1
                            }));
                            d3.selectAll('.thebar').classed('selected', false);
                        });
                    }
                }).style('margin-left', `${ratio * rowWidth}px`);

            d3.selectAll(thisObj.dot + row + index).append('p')
                .classed('valueText', true)
                // tslint:disable-next-line:no-any
                .each(function (datum: any, indexInner: number): void {
                    if (thisObj.colorByColumn !== -1) {
                        sum = thisObj.getSum(datum);
                    } else {
                        sum = datum[`values`][`value`];
                    }
                    const barwidth: number = sum <= 0 ? 0 :
                        (sum / thisObj.maxValue * (
                            (thisObj.binColumn === -1 ? 0.85 : 0.60) * (1 - ratio) * rowWidth));
                    const labelWidth: number = ratio * rowWidth;

                    let displayVal: number;
                    let tempMeasureData: string = sum.toString();
                    const valLen: number = tempMeasureData.split('.')[0].length;
                    displayVal = thisObj.returnDisplayVal(valLen);

                    //Value Formatter Creation
                    const formatter: IValueFormatter[] = [];
                    for (let iterator: number = 0; iterator < thisObj.dataView.metadata.columns.length; iterator++) {
                        const temp: IValueFormatter = ValueFormatter.create({
                            format: thisObj.dataView.metadata.columns[iterator].format,
                            value: displayVal
                        });
                        formatter.push(temp);
                    }

                    tempMeasureData = formatter[thisObj.targetColumn].format(sum.toFixed(2));
                    const textProperties: TextProperties = {
                        text: tempMeasureData,
                        fontFamily: thisObj.settings.fontSettings.fontFamily,
                        fontSize: `${thisObj.settings.fontSettings.fontSize}px`
                    };

                    const twidth: number = TextMeasurementService.measureSvgTextWidth(textProperties);
                    d3.select(this)
                        .style('left', sum <= 0 ? '5px' : (thisObj.value === 'outside' ?
                            `${barwidth + labelWidth - 55 + 5}px` :
                            ((twidth > barwidth) ? barwidth : `${barwidth + labelWidth - 55 - twidth - 2}px`)));
                })
                .style({
                    'font-size': `${this.settings.fontSettings.fontSize}px`,
                    'font-family': this.settings.fontSettings.fontFamily,
                    top: `${10 - thisObj.textSize / 2}px`,
                    //tslint:disable-next-line:no-any
                    'max-width': function (datum: any, indexInner: number): string {
                        const barwidth: number = sum <= 0 ? 0 :
                            (sum / thisObj.maxValue * (
                                (thisObj.binColumn === -1 ? 0.85 : 0.60) * (1 - ratio) * rowWidth));

                        return (thisObj.settings.value.displayValue === 'outside' ?
                            `${(thisObj.binColumn === -1 ? 0.15 : 0.40) * (1 - ratio) * rowWidth - 10}px` : `${barwidth}px`);
                        // 5 is padding + margin of label text + 5 is space between valuetext and bar
                    },
                    'white-space': 'nowrap',
                    overflow: 'hidden',
                    'text-overflow': 'ellipsis'
                })
                .style('color', thisObj.colors[thisObj.iColumntext][thisObj.jColumntext])
                // tslint:disable-next-line:no-any
                .text(function (datum: any): any {
                    if (thisObj.colorByColumn !== -1) {
                        sum = thisObj.getSum(datum);
                    } else {
                        sum = datum[`values`][`value`];
                    }
                    if (sum === 0) {
                        return `0`;
                    }
                    // tslint:disable-next-line:no-any
                    let displayVal: any;
                    let tempMeasureData: string = sum.toString();
                    const valLen: number = tempMeasureData.split('.')[0].length;
                    displayVal = thisObj.returnDisplayVal(valLen);

                    //Value Formatter Creation
                    const formatter: IValueFormatter[] = [];
                    for (let iterator: number = 0; iterator < thisObj.dataView.metadata.columns.length; iterator++) {
                        const temp: IValueFormatter = ValueFormatter.create({
                            format: thisObj.dataView.metadata.columns[iterator].format,
                            value: displayVal
                        });
                        formatter.push(temp);
                    }

                    tempMeasureData = formatter[thisObj.targetColumn].format(sum.toFixed(2));
                    const val: number = sum / thisObj.maxValue * (
                        (thisObj.binColumn === -1 ? 0.85 : 0.60) * chartWidth);
                    if (thisObj.value === 'inside' && val < 40) {
                        return '';
                    }

                    return tempMeasureData;
                })
                // tslint:disable-next-line:no-any
                .attr('title', function (datum: any): any {
                    if (thisObj.colorByColumn !== -1) {
                        sum = thisObj.getSum(datum);
                    } else {
                        sum = datum[`values`][`value`];
                    }

                    // tslint:disable-next-line:no-any
                    let displayVal: any;
                    let tempMeasureData: string = sum.toString();
                    const valLen: number = tempMeasureData.split('.')[0].length;
                    displayVal = thisObj.returnDisplayVal(valLen);

                    //Value Formatter Creation
                    const formatter: IValueFormatter[] = [];
                    for (let iterator: number = 0; iterator < thisObj.dataView.metadata.columns.length; iterator++) {
                        const temp: IValueFormatter = ValueFormatter.create({
                            format: thisObj.dataView.metadata.columns[iterator].format,
                            value: displayVal
                        });
                        formatter.push(temp);
                    }

                    tempMeasureData = formatter[thisObj.targetColumn].format(sum.toFixed(2));

                    if (sum === 0) {
                        return `0`;
                    }
                    const val: number = sum / thisObj.maxValue * (
                        (thisObj.binColumn === -1 ? 0.85 : 0.60) * chartWidth);
                    if (thisObj.value === 'inside' && val < 25) {
                        return '';
                    }

                    return tempMeasureData;
                });

            thisObj.renderLines('bar', subDiv, index, data, chartWidth, thisObj.maxValue, thisObj.minValue, thisObj.averageValue);

            //On Load Animation
            d3.selectAll('.chart .row')
                .style('transform-origin', '0% 0%')
                .style('transform', 'scale(0)')
                .transition()
                .duration(2000)
                .style('transform', 'scale(1)');
            chart.style('transform', 'scale(0)')
                .transition()
                .duration(1)
                .style('transform', 'scale(1)');
        }
        private ValueFormatter(valLen: number, tempMeasureData: number): string {
            let displayVal: string;
            const billion: string = 'B';
            const million: string = 'M';
            const thousand: string = 'K';
            if (valLen > 13) {
                displayVal = tempMeasureData.toString().substr(0, 3) + million;
            } else if (valLen === 12) {
                displayVal = tempMeasureData.toString().substr(0, 2) + million;
            } else if (valLen === 11) {
                displayVal = tempMeasureData.toString().substr(0, 1) + billion;
            } else if (valLen === 10) {
                displayVal = tempMeasureData.toString().substr(0, 3) + million;
            } else if (valLen === 9) {
                displayVal = tempMeasureData.toString().substr(0, 2) + million;
            } else if (valLen === 8) {
                displayVal = tempMeasureData.toString().substr(0, 1) + million;
            } else if (valLen === 7) {
                displayVal = tempMeasureData.toString().substr(0, 3) + thousand;
            } else if (valLen === 6) {
                displayVal = tempMeasureData.toString().substr(0, 2) + thousand;
            } else if (valLen === 5) {
                displayVal = tempMeasureData.toString().substr(0, 1) + thousand;
            } else {
                displayVal = tempMeasureData.toString();
            }

            return displayVal;
        }

        // tslint:disable-next-line:no-any
        private getMaxValue(data: any): void {
            const thisObj: this = this;
            let maxValue: number;
            let sum: number = 0;
            let minValue: number;

            minValue = thisObj.getSum(data[0]);
            maxValue = thisObj.getSum(data[0]);

            // tslint:disable-next-line:no-any
            data.forEach(function (d: number, i: number): any {
                if (thisObj.getSum(d) > maxValue) {
                    maxValue = thisObj.getSum(d);
                }
                if (thisObj.getSum(d) < minValue) {
                    minValue = thisObj.getSum(d);
                }
                sum = sum + thisObj.getSum(d);
            });
            this.averageValue = sum / data.length;
            this.maxValue = maxValue;
            this.minValue = minValue;
        }

        // tslint:disable-next-line:no-any
        private getMaxValue1(data: any): void {
            let iCounter: number;
            let maxValue: number;
            maxValue = 0;
            let sum: number;
            sum = 0;
            let minValue: number;
            minValue = data[0].values.value;
            for (iCounter = 0; iCounter < data.length; iCounter++) {
                if (data[iCounter].values.value > maxValue) {
                    maxValue = data[iCounter].values.value;
                }
                if (data[iCounter].values.value < minValue) {
                    minValue = data[iCounter].values.value;
                }
                sum = sum + data[iCounter].values.value;

            }
            this.averageValue = sum / data.length;
            this.maxValue = maxValue;
            this.minValue = minValue;
        }
        // tslint:disable-next-line:no-any
        private getSum(data: any): number {
            let sum: number = 0;
            const max: number = 0;
            // tslint:disable-next-line:no-any
            data.values[0].forEach(function (d: any, i: number): any {
                sum = sum + d[`values`][`value`];
            });

            return sum;
        }
        // tslint:disable-next-line:no-any
        private renderBarChart(data: any, index: number, categories: any, chartWidth: number,
                               chartHeight: number, height: number, marginLeft: number): void {
            this.noOfBars = data.length;
            const thisObj: this = this;
            thisObj.isCategory = true;
            const indexString: string = 'index';
            // tslint:disable-next-line:no-any
            let subDiv1: any;
            // tslint:disable-next-line:no-any
            let subDiv: any;
            // tslint:disable-next-line:no-any
            let chart1: any;
            const len: number = data.length;
            const categoriesLength: number = categories.length;
            const labelHeight: number = 16;
            const labelMarginBottom: number = 10;
            const marginBottom: number = 10;
            const marginBetweenBars: number = 1;
            const percentConverter: number = 100;
            // Ratio
            const ratio: number = thisObj.settings.ratio.percent / percentConverter;
            let rowWidth: number;
            if (thisObj.colorByColumn !== -1) {
                thisObj.getMaxValue(data);
            } else {
                thisObj.getMaxValue1(data);
            }
            // tslint:disable-next-line:no-any
            let chart: any;

            if (thisObj.binColumn === -1) {
                chart = thisObj.mainCont.append(`div`).classed('chart', true);
                subDiv = chart.append('div')
                    .classed('subDiv', true)
                    .classed('innerdiv', true)
                    .style({
                        width: $('.chart').width() - 20 + thisObj.px,
                        // tslint:disable-next-line:max-line-length
                        height: (Math.max((data.length * 15) + (marginBetweenBars * data.length) + labelHeight + labelMarginBottom + marginBottom, height)) + thisObj.px,
                        padding: '0',
                        margin: '0'
                    });

            } else {
                chart = thisObj.mainCont
                    .append('div')
                    .attr('class', 'chart borderChart binChart')
                    .style({
                        position: 'relative',
                        width: chartWidth + thisObj.px,
                        height: chartHeight + thisObj.px,
                        'margin-left': marginLeft + thisObj.px,
                        overflow: 'hidden'
                    });
                subDiv1 = chart.append('div')
                    .style({
                        'min-height': '10px',
                        width: '100%',
                        padding: '0px',
                        margin: '0px',
                        'text-align': 'center',
                        'font-weight': 'bold',
                        height: '20px'
                    });

                subDiv1.append('label')
                    .text(thisObj.returnLabelText(categories[index]) || 'Blank')
                    .attr('title', thisObj.returnLabelText(categories[index]))
                    .classed('label', true);

                chart1 = chart.append('div')
                    .style({
                        position: 'relative',
                        'max-height': '100%',
                        overflow: 'auto'
                    });

                subDiv = chart1.append('div')
                    .classed('subDiv', true)
                    .classed('innerdiv', true)
                    .style({
                        width: chartWidth + thisObj.px,
                        'min-height': chartHeight + thisObj.px,
                        padding: '0',
                        margin: '0',
                        overflow: 'auto'

                    });

                subDiv.append('label');
            }
            let percentSign: string;
            percentSign = '%';

            let increment: number;
            increment = thisObj.previousDataLength + thisObj.colorByDataSelection;
            const row: string = 'row';
            // tslint:disable-next-line:no-any
            let tooltipchart: any;
            tooltipchart = subDiv.selectAll('div')
                .data(data)
                .enter()
                .append('div')
                .classed(row + index, true)
                .classed('row', true);

            rowWidth = parseInt(d3.select('.row').style('width').split('p')[0], 10) +
                parseInt(d3.select('.row').style('margin-left').split('p')[0], 10);

            d3.selectAll(thisObj.dot + row + index)
                .append('p')
                .classed('categoryText', true)
                // tslint:disable-next-line:no-any
                .attr('title', function (datum: any): string { return datum[`key`].toString(); })
                // tslint:disable-next-line:no-any
                .style('margin-bottom', function (datum: any, iterator: number): string {
                    if (iterator === data.length - 1) {
                        return marginBottom + thisObj.px;
                    }

                    return '1px';
                })
                // tslint:disable-next-line:no-any
                .text(function (datum: any): string {
                    let labelText: string;
                    labelText = datum[`key`].toString() || 'Blank';

                    return labelText;
                });

            let sum: number;

            // tslint:disable-next-line:no-any
            const tooltipData: any = [];
            d3.selectAll(thisObj.dot + row + index)
                .append('div')
                .classed('thebar', true)
                // tslint:disable-next-line:no-any
                .each(function (datum: any, iIterator: number): any {
                    // tslint:disable-next-line:no-any
                    if (thisObj.colorByColumn !== -1) {
                        // tslint:disable-next-line:no-any
                        const THIS: any = this;
                        sum = thisObj.getSum(datum);
                        // tslint:disable-next-line:no-any
                        datum.values[0].forEach(function (d: any, i: number): any {
                            thisObj.globalSelections.push({
                                data: datum,
                                data1: datum.values[0],
                                binIndex: index,
                                category: categories[index]
                            });

                            const value: number = d[`values`][`value`];

                            if (value > 0) {
                                thisObj.isCategory = false;
                            }
                            thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                     (tooltipEvent: TooltipEventArgs<number>) =>
                                    thisObj.getTooltipData(tooltipEvent.data, d, 'bar', '', 0, thisObj),
                                                                     (tooltipEvent: TooltipEventArgs<number>) => null);
                            const p: number = d[`values`][`value`] / sum;
                            tooltipData.push({ key: d[`key`], value: d[`values`][`value`] });
                            // thisObj.selectionManager.clear();
                            if (thisObj.colorByColumn !== thisObj.binColumn) {
                                d3.select(THIS)
                                    .append('div').classed('inDiv', true)
                                    .style('float', 'left')
                                    .style('margin', '0')
                                    .style('width', d[`values`][`value`] <= 0 ? `0` :
                                        // tslint:disable-next-line:max-line-length
                                        `${p * ((sum / thisObj.maxValue) * (thisObj.binColumn === -1 ? 0.85 : 0.60) * (1 - ratio) * rowWidth)}px`)
                                    // tslint:disable-next-line:no-any
                                    .on('click', function (dat: any): void {
                                        const selectionlen: number = dat.values[0].length;
                                        let level: number = 0;
                                        for (i = 0; i < selectionlen; i++) {
                                            if (dat.values[0][i].key === d.key) {
                                                level = i;
                                                break;
                                            }
                                        }
                                        // tslint:disable-next-line:no-any
                                        const selectionIds: any[] = dat.values[0][level].values.selectionId;
                                        // tslint:disable-next-line:no-any
                                        const lastindex: any = thisObj.selectionManager;
                                        // tslint:disable-next-line:no-any
                                        const lastIndexLen: any = lastindex.selectedIds.length;
                                        const isCrtlPressed: boolean = (d3.event as MouseEvent).ctrlKey;
                                        thisObj.selectionManager.select(selectionIds, isCrtlPressed).then((ids: ISelectionId[]) => {
                                            if (isCrtlPressed) {
                                                if (lastIndexLen === 0 && !(d3.select(this).classed('selected'))) {
                                                    d3.selectAll('.inDiv').style({
                                                        opacity: 0.3
                                                    });
                                                    d3.select(this).style({
                                                        opacity: 1
                                                    })
                                                    .classed('selected', true);
                                                    thisObj.lastSelectedbar = 1;
                                                } else if (d3.select(this).classed('selected')) {
                                                    d3.select(this).style({
                                                        opacity: 0.3
                                                    })
                                                    .classed('selected', false);
                                                    thisObj.lastSelectedbar--;
                                                    if (thisObj.lastSelectedbar === 0) {
                                                        d3.selectAll('.inDiv').style({
                                                            opacity: 1
                                                        });
                                                    }
                                                } else {
                                                    d3.select(this).style({
                                                        opacity: 1
                                                    })
                                                    .classed('selected', true);
                                                    thisObj.lastSelectedbar++;
                                                }
                                            } else {
                                                if (d3.select(this).classed('selected')) {
                                                    d3.select(this).classed('selected', false);
                                                    thisObj.selectionManager.clear().then(() => d3.selectAll('.inDiv').style({
                                                        opacity: 1
                                                    }));
                                                    if (!isCrtlPressed) {
                                                        thisObj.lastSelectedbar--;
                                                    }
                                                } else {
                                                d3.selectAll('.inDiv').style({
                                                    opacity: 0.3
                                                })
                                                .classed('selected', false);
                                                d3.select(this).style({
                                                    opacity: 1
                                                })
                                                .classed('selected', true);
                                                thisObj.lastSelectedbar = 0;
                                                if (!isCrtlPressed) {
                                                    thisObj.lastSelectedbar++;
                                                }
                                            }
                                            }
                                        });
                                        (<Event>d3.event).stopPropagation();
                                    })
                                    .each(function (): void {
                                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                                thisObj.getTooltipData(tooltipEvent.data, d, 'bar', '', 0, thisObj),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);
                                    })
                                    // tslint:disable-next-line:no-any
                                    .attr('data-selection', function (datumm: any, iterator: number): number { return increment; })
                                    // tslint:disable-next-line:no-any
                                    .style('background-color', function (): any {
                                        // tslint:disable-next-line:no-any
                                        return thisObj.barcolor.filter(function (v: any): any {
                                            return v.key.toString() === d[`key`].toString();
                                        })[0].value;
                                    });
                                // tslint:disable-next-line:typedef
                                subDiv.on('click', function () {
                                    thisObj.selectionManager.clear().then(() => d3.selectAll('.inDiv').style({
                                        opacity: 1
                                    }));
                                    d3.selectAll('.inDiv').classed('selected', false);
                                });
                                increment++;
                            } else {
                                d3.select(THIS)
                                    .append('div').classed('inDiv', true)
                                    .style('float', 'left')
                                    .style('margin', '0')
                                    .style('width', d[`values`][`value`] <= 0 ? `0` :
                                        // tslint:disable-next-line:max-line-length
                                        `${p * ((sum / thisObj.maxValue) * (thisObj.binColumn === -1 ? 0.85 : 0.60) * (1 - ratio) * rowWidth)}px`)
                                    // tslint:disable-next-line:no-any
                                    .on('click', function (dat: any): void {
                                        const selectionlen: number = dat.values[0].length;
                                        let level: number = 0;
                                        for (i = 0; i < selectionlen; i++) {
                                            if (dat.values[0][i].key === d.key) {
                                                level = i;
                                                break;
                                            }
                                        }
                                        // tslint:disable-next-line:no-any
                                        const selectionIds: any[] = dat.values[0][level].values.selectionId;
                                        // tslint:disable-next-line:no-any
                                        const lastindex: any = thisObj.selectionManager;
                                        // tslint:disable-next-line:no-any
                                        const lastIndexLen: any = lastindex.selectedIds.length;
                                        const isCrtlPressed: boolean = (d3.event as MouseEvent).ctrlKey;
                                        thisObj.selectionManager.select(selectionIds, isCrtlPressed).then((ids: ISelectionId[]) => {
                                            if (isCrtlPressed) {
                                                if (lastIndexLen === 0 && !(d3.select(this).classed('selected'))) {
                                                    d3.selectAll('.inDiv').style({
                                                        opacity: 0.3
                                                    });
                                                    d3.select(this).style({
                                                        opacity: 1
                                                    })
                                                    .classed('selected', true);
                                                    thisObj.lastSelectedbar = 1;
                                                } else if (d3.select(this).classed('selected')) {
                                                    d3.select(this).style({
                                                        opacity: 0.3
                                                    })
                                                    .classed('selected', false);
                                                    thisObj.lastSelectedbar--;
                                                    if (thisObj.lastSelectedbar === 0) {
                                                        d3.selectAll('.inDiv').style({
                                                            opacity: 1
                                                        });
                                                    }
                                                } else {
                                                    d3.select(this).style({
                                                        opacity: 1
                                                    })
                                                    .classed('selected', true);
                                                    thisObj.lastSelectedbar++;
                                                }
                                            } else {
                                                if (d3.select(this).classed('selected')) {
                                                    d3.select(this).classed('selected', false);
                                                    thisObj.selectionManager.clear().then(() => d3.selectAll('.inDiv').style({
                                                        opacity: 1
                                                    }));
                                                    if (!isCrtlPressed) {
                                                        thisObj.lastSelectedbar--;
                                                    }
                                                } else {
                                                d3.selectAll('.inDiv').style({
                                                    opacity: 0.3
                                                })
                                                .classed('selected', false);
                                                d3.select(this).style({
                                                    opacity: 1
                                                })
                                                .classed('selected', true);
                                                thisObj.lastSelectedbar = 0;
                                                if (!isCrtlPressed) {
                                                    thisObj.lastSelectedbar++;
                                                }
                                            }
                                            }
                                        });
                                        (<Event>d3.event).stopPropagation();
                                    })
                                    .each(function (): void {
                                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                                thisObj.getTooltipData(tooltipEvent.data, d, 'bar', '', 0, thisObj),
                                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);
                                    })
                                    // tslint:disable-next-line:no-any
                                    .attr('data-selection', function (datumm: any, iterator: number): number { return increment; })
                                    // tslint:disable-next-line:no-any
                                    .style('background-color', thisObj.barcolor.filter(function (v: any): any {
                                        return v.key.toString() === d[`key`].toString();
                                    })[0].value);
                                // tslint:disable-next-line:typedef
                                subDiv.on('click', function () {
                                    thisObj.selectionManager.clear().then(() => d3.selectAll('.inDiv').style({
                                        opacity: 1
                                    }));
                                    d3.selectAll('.inDiv').classed('selected', false);
                                });
                                increment++;
                            }
                        });

                    } else {
                        // tslint:disable-next-line:no-any
                        const value: any = datum[`values`][`value`];
                        if (value > 0) {
                            thisObj.isCategory = false;
                        }
                        d3.select(this)
                            .style('width', datum[`values`][`value`] <= 0 ? `0` :
                                `${(parseInt(datum[`values`][`value`], 10) / thisObj.maxValue * (
                                    (thisObj.binColumn === -1 ? 0.85 : 0.60) * (1 - ratio) * rowWidth))}px`)
                            .style('background-color', thisObj.colors[thisObj.iColumn][thisObj.jColumn])
                            // tslint:disable-next-line:no-any
                            .on('click', function (d: any): void {
                                // tslint:disable-next-line:no-any
                                const selectionIds: any[] = datum[`values`][`selectionId`];
                                // tslint:disable-next-line:no-any
                                const lastindex: any = thisObj.selectionManager;
                                // tslint:disable-next-line:no-any
                                const lastIndexLen: any = lastindex.selectedIds.length;
                                const isCrtlPressed: boolean = (d3.event as MouseEvent).ctrlKey;
                                thisObj.selectionManager.select(selectionIds, isCrtlPressed)
                                    .then((ids: ISelectionId[]) => {
                                        if (isCrtlPressed) {
                                            if (lastIndexLen === 0 && !(d3.select(this).classed('selected'))) {
                                                d3.selectAll('.thebar').style({
                                                    opacity: 0.3
                                                });
                                                d3.select(this).style({
                                                    opacity: 1
                                                })
                                                .classed('selected', true);
                                                thisObj.lastSelectedbar = 1;
                                            } else if (d3.select(this).classed('selected')) {
                                                d3.select(this).style({
                                                    opacity: 0.3
                                                })
                                                .classed('selected', false);
                                                thisObj.lastSelectedbar--;
                                                if (thisObj.lastSelectedbar === 0) {
                                                    d3.selectAll('.thebar').style({
                                                        opacity: 1
                                                    });
                                                }
                                            } else {
                                                d3.select(this).style({
                                                    opacity: 1
                                                })
                                                .classed('selected', true);
                                                thisObj.lastSelectedbar++;
                                            }
                                        } else {
                                            if (d3.select(this).classed('selected')) {
                                                d3.select(this).classed('selected', false);
                                                thisObj.selectionManager.clear().then(() => d3.selectAll('.thebar').style({
                                                    opacity: 1
                                                }));
                                                if (!isCrtlPressed) {
                                                    thisObj.lastSelectedbar--;
                                                }
                                            } else {
                                            d3.selectAll('.thebar').style({
                                                opacity: 0.3
                                            })
                                            .classed('selected', false);
                                            d3.select(this).style({
                                                opacity: 1
                                            })
                                            .classed('selected', true);
                                            thisObj.lastSelectedbar = 0;
                                            if (!isCrtlPressed) {
                                                thisObj.lastSelectedbar++;
                                            }
                                        }
                                        }
                                    });
                                (<Event>d3.event).stopPropagation();
                            });
                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                thisObj.getTooltipData(tooltipEvent.data, datum, 'bin', '', 0, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);
                        // tslint:disable-next-line:typedef
                        subDiv.on('click', function () {
                            thisObj.selectionManager.clear().then(() => d3.selectAll('.thebar').style({
                                opacity: 1
                            }));
                            d3.selectAll('.thebar').classed('selected', false);
                        });
                    }
                })
                .style('margin-left', `${ratio * rowWidth}px`);
            thisObj.previousDataLength = increment;
            // tslint:disable-next-line:no-any
            d3.selectAll(thisObj.dot + row + index).append('p')
                .classed('valueText', true)
                // tslint:disable-next-line:no-any
                .each(function (datum: any, indexInner: number): void {
                    if (thisObj.colorByColumn !== -1) {
                        sum = thisObj.getSum(datum);
                    } else {
                        sum = datum[`values`][`value`];
                    }
                    const barwidth: number = sum <= 0 ? 0 :
                        (sum / thisObj.maxValue * (
                            (thisObj.binColumn === -1 ? 0.85 : 0.60) * (1 - ratio) * rowWidth));
                    const labelWidth: number = ratio * rowWidth;

                    let displayVal: number;
                    let tempMeasureData: string = sum.toString();
                    const valLen: number = tempMeasureData.split('.')[0].length;
                    displayVal = thisObj.returnDisplayVal(valLen);

                    //Value Formatter Creation
                    const formatter: IValueFormatter[] = [];
                    for (let iterator: number = 0; iterator < thisObj.dataView.metadata.columns.length; iterator++) {
                        const temp: IValueFormatter = ValueFormatter.create({
                            format: thisObj.dataView.metadata.columns[iterator].format,
                            value: displayVal
                        });
                        formatter.push(temp);
                    }

                    tempMeasureData = formatter[thisObj.targetColumn].format(sum.toFixed(2));
                    const textProperties: TextProperties = {
                        text: tempMeasureData,
                        fontFamily: thisObj.settings.fontSettings.fontFamily,
                        fontSize: `${thisObj.settings.fontSettings.fontSize}px`
                    };

                    const twidth: number = TextMeasurementService.measureSvgTextWidth(textProperties);
                    d3.select(this)
                        .style('left', sum <= 0 ? '5px' : (thisObj.value === 'outside' ?
                            `${barwidth + labelWidth - 55 + 5}px` :
                            ((twidth > barwidth) ? barwidth : `${barwidth + labelWidth - 55 - twidth - 2}px`)));
                })
                .style({
                    'font-size': `${this.settings.fontSettings.fontSize}px`,
                    'font-family': this.settings.fontSettings.fontFamily,
                    top: `${10 - thisObj.textSize / 2}px`,
                    //tslint:disable-next-line:no-any
                    'max-width': function (datum: any, indexInner: number): string {
                        const barwidth: number = sum <= 0 ? 0 :
                            (sum / thisObj.maxValue * (
                                (thisObj.binColumn === -1 ? 0.85 : 0.60) * (1 - ratio) * rowWidth));

                        return (thisObj.settings.value.displayValue === 'outside' ?
                            `${(thisObj.binColumn === -1 ? 0.15 : 0.40) * (1 - ratio) * rowWidth - 10}px` : `${barwidth}px`);
                        // 5 is padding + margin of label text + 5 is space between valuetext and bar
                    },
                    'white-space': 'nowrap',
                    overflow: 'hidden',
                    'text-overflow': 'ellipsis'
                })
                .style('color', thisObj.colors[thisObj.iColumntext][thisObj.jColumntext])
                // tslint:disable-next-line:no-any
                .text(function (datum: any): any {
                    if (thisObj.colorByColumn !== -1) {
                        sum = thisObj.getSum(datum);
                    } else {
                        sum = datum[`values`][`value`];
                    }
                    if (sum === 0) {
                        return `0`;
                    }
                    // tslint:disable-next-line:no-any
                    let displayVal: any;
                    let tempMeasureData: string = sum.toString();
                    const valLen: number = tempMeasureData.split('.')[0].length;
                    displayVal = thisObj.returnDisplayVal(valLen);

                    //Value Formatter Creation
                    const formatter: IValueFormatter[] = [];
                    for (let iterator: number = 0; iterator < thisObj.dataView.metadata.columns.length; iterator++) {
                        const temp: IValueFormatter = ValueFormatter.create({
                            format: thisObj.dataView.metadata.columns[iterator].format,
                            value: displayVal
                        });
                        formatter.push(temp);
                    }

                    tempMeasureData = formatter[thisObj.targetColumn].format(sum.toFixed(2));
                    const val: number = sum / thisObj.maxValue * (
                        (thisObj.binColumn === -1 ? 0.85 : 0.60) * chartWidth);
                    if (thisObj.value === 'inside' && val < 40) {
                        return '';
                    }

                    return tempMeasureData;
                })
                // tslint:disable-next-line:no-any
                .attr('title', function (datum: any): any {
                    if (thisObj.colorByColumn !== -1) {
                        sum = thisObj.getSum(datum);
                    } else {
                        sum = datum[`values`][`value`];
                    }

                    // tslint:disable-next-line:no-any
                    let displayVal: any;
                    let tempMeasureData: string = sum.toString();
                    const valLen: number = tempMeasureData.split('.')[0].length;
                    displayVal = thisObj.returnDisplayVal(valLen);

                    //Value Formatter Creation
                    const formatter: IValueFormatter[] = [];
                    for (let iterator: number = 0; iterator < thisObj.dataView.metadata.columns.length; iterator++) {
                        const temp: IValueFormatter = ValueFormatter.create({
                            format: thisObj.dataView.metadata.columns[iterator].format,
                            value: displayVal
                        });
                        formatter.push(temp);
                    }

                    tempMeasureData = formatter[thisObj.targetColumn].format(sum.toFixed(2));

                    if (sum === 0) {
                        return `0`;
                    }
                    const val: number = sum / thisObj.maxValue * (
                        (thisObj.binColumn === -1 ? 0.85 : 0.60) * chartWidth);
                    if (thisObj.value === 'inside' && val < 25) {
                        return '';
                    }

                    return tempMeasureData;
                });
            thisObj.renderLines('bar', subDiv, index, data, chartWidth, thisObj.maxValue, thisObj.minValue, thisObj.averageValue);
            //On Load Animation
            d3.selectAll('.chart .row')
                .style('transform-origin', '0% 0%')
                .style('transform', 'scale(0)')
                .transition()
                .duration(2000)
                .style('transform', 'scale(1)');

            chart.style('transform', 'scale(0)')
                .transition()
                .duration(1)
                .style('transform', 'scale(1)');
        }

        private renderStyles(elementClass: string, styleDivClass: string, stylePClass: string,
                             lineStyle: string, lineFill: string, lineOpacity: number, lineValue: number): void {
            const thisObj: this = this;
            d3.select(thisObj.dot + elementClass).on('mouseover', () => {
                d3.select(thisObj.dot + elementClass)
                    // on hover line should be hightlighted hence 3
                    .style(elementClass.substr(0, 6) === 'column' ? 'border-top' : 'border-right',
                           3 + thisObj.px + thisObj.space + lineStyle + thisObj.space + lineFill);
            });
            d3.select(thisObj.dot + elementClass).on('mouseout', () => {
                d3.select(thisObj.dot + elementClass)
                    .style(elementClass.substr(0, 6) === 'column' ? 'border-top' : 'border-right',
                           1 + thisObj.px + thisObj.space + lineStyle +
                        thisObj.space + lineFill);
            });

            d3.selectAll(thisObj.dot + styleDivClass)
                .style(elementClass.substr(0, 6) === 'column' ? 'border-top' : 'border-right',
                       1 + thisObj.px + thisObj.space + lineStyle + thisObj.space + lineFill)
                .style('opacity', lineOpacity / 100)
                .style('position', 'absolute');

            d3.selectAll(thisObj.dot + stylePClass)
                .style('color', lineFill)
                .style('opacity', lineOpacity / 100);
        }

        private percentOfBar(): number {
            if (this.binColumn === -1) {
                return 0.85;
            }

            return 0.60;
        }

        private returnChartWidth(chartWidth: number): number {
            if (this.binColumn === -1) {
                return 0;
            } else {
                return chartWidth;
            }
        }
        private returnAdditionMargin(): number {
            if (this.binColumn === -1) {
                return -3;
            }

            return 0;
        }

        private returnLabelText(text: string): string {
            if (this.binColumn === -1) {
                return '';
            }
            if (text === null) {
                return 'Null';
            }

            return this.binColumnformatter.format(text);
        }
        private returnDisplayVal(valLen: number): number {
            if (valLen > 9) {
                return 1e9;
            } else if (valLen <= 9 && valLen > 6) {
                return 1e6;
            } else if (valLen <= 6 && valLen >= 4) {
                return 1e3;
            } else {
                return 10;
            }
        }

        private returnMarginForCenter(indexCounter: number, skip: number, chartWidth: number): number {
            return indexCounter % skip === 0 ? (this.mainContWidth - (chartWidth * skip) - (10 * skip)) / 2 : 5;

        }

        // tslint:disable-next-line:no-any
        private returnHeight(categories: any, chartHeight: number): number {
            if (categories >= 3) {
                return this.mainContHeight - 50;
            }

            return chartHeight;
        }

        private returnRenderHeight(subDivHeight: number, labelHeight: number,
                                   labelMarginBottom: number, colValueText: number, categoryHeight: number): number {
            if (this.binColumn === -1) {
                return this.mainContHeight - 30 - labelHeight - labelMarginBottom - colValueText - categoryHeight;
            }
            // here 20 ponits of length is decreased from renderHeight to provide extra space on increasing font

            return subDivHeight - labelHeight - labelMarginBottom  - colValueText - categoryHeight;
        }
        // tslint:disable-next-line:no-any
        private renderLines(chartType: string, chart: any, index: number, data: any,
                            chartWidth: number, maxValue: number, minValue: number, averageValue: number): void {
            const thisObj: this = this;
            const extraMargin: number = 5;
            const row: string = 'row';
            // tslint:disable-next-line:no-any
            const $label: any = d3.select('.label');
            // tslint:disable-next-line:no-any
            const bIsLabelVal: any = null === $($label[0])[0] ? 1 : 0;
            let labelMarginBottom: number = 0;
            let labelPaddingTop: number = 0;
            let labelPaddingBottom: number = 0;
            let labelHeight: number = 16;
            let labelMarginTop: number = 0;
            // tslint:disable-next-line:no-any
            const $barDiv: any = d3.select('.thebar');
            // tslint:disable-next-line:no-any
            const $currentDiv: any = d3.select(thisObj.dot + row + index);
            if (!bIsLabelVal) {
                labelMarginBottom = parseInt(($label.style('margin-bottom').substr(0, $label.style('margin-bottom').length - 2)), 10);
                labelPaddingTop = parseInt(($label.style('padding-top').substr(0, $label.style('padding-top').length - 2)), 10);
                labelPaddingBottom = parseInt(($label.style('padding-bottom').substr(0, $label.style('padding-bottom').length - 2)), 10);
                labelHeight = $('.label').height();
                labelMarginTop = parseInt(($label.style('margin-top').substr(0, $label.style('margin-top').length - 2)), 10);
            }
            let thebarHeight: number = $(thisObj.dot + row + index).height();
            thebarHeight = 35;
            const thebarMarginTop: number =
                parseInt(($currentDiv.style('margin-top').substr(0, $barDiv.style('margin-top').length - 2)), 10);
            const thebarMarginBottom: number =
                parseInt(($currentDiv.style('margin-bottom').substr(0, $barDiv.style('margin-bottom').length - 2)), 10);
            const thebarPaddingBottom: number =
                parseInt(($currentDiv.style('padding-bottom').substr(0, $barDiv.style('padding-bottom').length - 2)), 10);
            const thebarPaddingtop: number =
                parseInt(($currentDiv.style('padding-top').substr(0, $barDiv.style('padding-top').length - 2)), 10);
            const top: string = thisObj.binColumn === -1 ? labelHeight + labelMarginBottom + 3 * thebarMarginTop + 10 + thisObj.px
                : labelHeight + labelMarginBottom + 10 + thisObj.px;
            const chartHeight: number =
                labelMarginTop + labelHeight + labelMarginBottom + labelPaddingTop + labelPaddingBottom + thebarMarginTop
                + thebarMarginBottom * data.length + thebarHeight * data.length + thebarPaddingBottom + thebarMarginTop * data.length;
            const rowWidth: number = parseInt(d3.select('.row').style('width').split('p')[0], 10) +
                parseInt(d3.select('.row').style('margin-left').split('p')[0], 10);
            const percentConverter: number = 100;
            const ratio: number = thisObj.settings.ratio.percent / percentConverter;

            const maxLine: string = 'maxLine';
            const divMaxLine: string = 'max';
            const pMaxLine: string = 'pMax';
            const divMinLine: string = 'min';
            const pMinLine: string = 'pMin';
            const divAvgLine: string = 'avg';
            const pAvgLine: string = 'pAvg';
            const divConstantLine: string = 'constant';
            const pConstantLine: string = 'pConstant';
            const barPercent: number = thisObj.percentOfBar();

            // MAX LINE
            if (thisObj.maxLineIsOn) {
                chart.append('div')
                    .style('margin-top', function (): string {
                        if (thisObj.binColumn === -1) {
                            return '-28px';
                        } else {
                            return '4px';
                        }
                    })
                    .style('left', thisObj.isCategory ? 0 : (ratio * rowWidth) + (barPercent * (1 - ratio) * rowWidth) + 5 + thisObj.px)
                    .classed(maxLine + index, true)
                    .classed(divMaxLine, true)
                    .style('margin-left', function (): string {
                        if (maxValue !== 0) {
                            return '0px';
                        } else {
                            return '60px';
                        }
                    })
                    // .style('margin-top', '5px')
                    .style('height', ((this.noOfBars * 15) + this.noOfBars - 2) + thisObj.px)
                    .style('top', function (): string {
                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                thisObj.getTooltipDataAnalyticsLine(tooltipEvent.data, '', 'line', 'Max', maxValue, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);

                        return top;
                    });
                if (thisObj.maxLineDataLabel) {
                    chart.append('p')
                        .classed(pMaxLine, true)
                        .style('margin-left', 0)
                        .style('left', function (): string {
                            if (maxValue !== 0) {
                                return (ratio * rowWidth) + (barPercent * (1 - ratio) * rowWidth) - 10 + thisObj.px;
                            } else {
                                return '35px';
                            }
                        })
                        .classed('linesP', true)
                        .attr('title', thisObj.maxValue)
                        .style('margin', '2px')
                        .text(function (): string {
                            let displayVal: number;
                            let tempMeasureData: string = thisObj.maxValue.toString();
                            const valLen: number = tempMeasureData.split('.')[0].length;
                            displayVal = thisObj.returnDisplayVal(valLen);

                            //Value Formatter Creation
                            const formatter: IValueFormatter = ValueFormatter.create({
                                value: displayVal
                            });

                            tempMeasureData = formatter.format(thisObj.maxValue);
                            $(this).attr('title', tempMeasureData);

                            return tempMeasureData;
                        })
                        .style({
                            'font-size': `${this.settings.fontSettings.fontSize}px`,
                            'font-family': this.settings.fontSettings.fontFamily,
                            top: `-${((thisObj.textSize / 4) - 25 + thisObj.textSize)}px`,
                            width: `${thisObj.textWidth}px`
                        });
                }
                thisObj.renderStyles(maxLine + index, divMaxLine, pMaxLine, thisObj.maxLineStyle,
                                     thisObj.maxLineFill, thisObj.maxLineOpacity, thisObj.maxValue);
            }
            // Render Min Line
            if (thisObj.minLineIsOn) {
                const minLine: string = 'minLine';
                chart.append('div')
                    .style('margin-top', function (): string {
                        if (thisObj.binColumn === -1) {
                            return '-28px';
                        } else {
                            return '4px';
                        }
                    })
                    .classed(minLine + index, true)
                    .classed(divMinLine, true)
                    .style('margin-left', function (): string {
                        if (maxValue !== 0) {
                            return '0px';
                        } else {
                            return '60px';
                        }
                    })
                    //.style('margin-top', '5px')
                    .style('left', thisObj.isCategory ? '0' : thisObj.minValue / thisObj.maxValue
                        * barPercent * (1 - ratio) * rowWidth + ratio * rowWidth + 5 + thisObj.px)
                    .style('height', ((this.noOfBars * 15) + this.noOfBars - 2) + thisObj.px)
                    .style('top', function (): string {
                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                thisObj.getTooltipDataAnalyticsLine(tooltipEvent.data, '', 'line', 'Min', minValue, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);

                        return top;
                    });
                if (thisObj.minLineDataLabel) {
                    chart.append('p')
                        .classed(pMinLine, true)
                        .style('margin-left', 0)
                        .style('left', function (): string {
                            if (maxValue !== 0) {
                                return thisObj.minValue / thisObj.maxValue
                                    * barPercent * (1 - ratio) * rowWidth + ratio * rowWidth - 19 + thisObj.px;
                            } else {
                                return '35px';
                            }
                        })
                        .classed('linesP', true)
                        .attr('title', thisObj.minValue)
                        .text(function (): string {
                            let displayVal: number;
                            let tempMeasureData: string = thisObj.minValue.toString();
                            const valLen: number = tempMeasureData.split('.')[0].length;
                            displayVal = thisObj.returnDisplayVal(valLen);

                            //Value Formatter Creation
                            const formatter: IValueFormatter = ValueFormatter.create({
                                value: displayVal
                            });

                            tempMeasureData = formatter.format(thisObj.minValue);
                            $(this).attr('title', tempMeasureData);

                            return tempMeasureData;
                        })
                        .style({
                            'font-size': `${this.settings.fontSettings.fontSize}px`,
                            'font-family': this.settings.fontSettings.fontFamily,
                            top: `-${((thisObj.textSize / 4) - 25 + thisObj.textSize)}px`,
                            width: `${thisObj.textWidth}px`
                        });
                }
                thisObj.renderStyles(minLine + index, divMinLine, pMinLine, thisObj.minLineStyle,
                                     thisObj.minLineFill, thisObj.minLineOpacity, thisObj.minValue);
            }
            // Render Average Line
            if (thisObj.avgLineIsOn) {
                const averageLine: string = 'averageLine';
                chart.append('div')
                    .style('margin-top', function (): string {
                        if (thisObj.binColumn === -1) {
                            return '-28px';
                        } else {
                            return '4px';
                        }
                    })
                    .classed(averageLine + index, true)
                    .classed(divAvgLine, true)
                    .style('margin-left', function (): string {
                        if (maxValue !== 0) {
                            return '0px';
                        } else {
                            return '60px';
                        }
                    })
                    //.style('margin-top', '5px')
                    .style('left', thisObj.isCategory ? 0 : thisObj.averageValue / thisObj.maxValue
                        * barPercent * (1 - ratio) * rowWidth + ratio * rowWidth + thisObj.px)

                    .style('height', ((this.noOfBars * 15) + this.noOfBars - 2) + thisObj.px)
                    .style('top', function (): string {
                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                // tslint:disable-next-line:max-line-length
                                thisObj.getTooltipDataAnalyticsLine(tooltipEvent.data, '', 'line', 'Average', thisObj.averageValue, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);

                        return top;
                    });
                if (thisObj.avgLineDataLabel) {
                    chart.append('p')
                        .classed(pAvgLine, true)
                        .style('margin-left', 0)
                        .style('left', function (): string {
                            if (maxValue !== 0) {
                                return thisObj.minValue / thisObj.maxValue
                                    * barPercent * (1 - ratio) * rowWidth + ratio * rowWidth + 12 + thisObj.px;
                            } else {
                                return '35px';
                            }
                        })
                        .classed('linesP', true)
                        .attr('title', thisObj.averageValue)
                        .text(function (): string {
                            let displayVal: number;
                            let tempMeasureData: string = thisObj.averageValue.toString();
                            const valLen: number = tempMeasureData.split('.')[0].length;
                            displayVal = thisObj.returnDisplayVal(valLen);

                            //Value Formatter Creation
                            const formatter: IValueFormatter = ValueFormatter.create({
                                value: displayVal
                            });

                            tempMeasureData = formatter.format(thisObj.averageValue);
                            $(this).attr('title', tempMeasureData);

                            return tempMeasureData;
                        })
                        .style({
                            'font-size': `${this.settings.fontSettings.fontSize}px`,
                            'font-family': this.settings.fontSettings.fontFamily,
                            top: `-${((thisObj.textSize / 4) - 25 + thisObj.textSize)}px`,
                            width: `${thisObj.textWidth}px`
                        });
                }
                thisObj.renderStyles(averageLine + index, divAvgLine, pAvgLine, thisObj.avgLineStyle,
                                     thisObj.avgLineFill, thisObj.avgLineOpacity, thisObj.averageValue);
            }
            // Render Constant Line
            if (thisObj.constantLineIsOn) {
                const constantLine: string = 'constantLine';
                chart.append('div')
                    .style('margin-top', function (): string {
                        if (thisObj.binColumn === -1) {
                            return '-28px';
                        } else {
                            return '4px';
                        }
                    })
                    .classed(divConstantLine, true)
                    .classed(constantLine + index, true)
                    .style('margin-left', function (): string {
                        if (maxValue !== 0) {
                            return '0px';
                        } else {
                            return '60px';
                        }
                    })
                    //.style('margin-top', '5px')
                    .style('left', thisObj.isCategory ? 0 : thisObj.constantLineValue / thisObj.maxValue
                        * barPercent * (1 - ratio) * rowWidth + ratio * rowWidth + 5 + thisObj.px)
                    .style('height', ((this.noOfBars * 15) + this.noOfBars - 2) + thisObj.px)
                    .style('top', top)
                    .each(function (): void {
                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                // tslint:disable-next-line:max-line-length
                                thisObj.getTooltipDataAnalyticsLine(tooltipEvent.data, '', 'line', 'Constant', thisObj.constantLineValue, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);
                    });
                if (thisObj.constantLineDataLabel) {
                    chart.append('p')
                        .classed(pConstantLine, true)
                        .style('margin-top', '7px')
                        .style('margin-left', '-30px')
                        .style('left', function (): string {
                            if (maxValue !== 0) {
                                return thisObj.isCategory ? '0'
                                    : thisObj.constantLineValue / thisObj.maxValue
                                    * barPercent * (1 - ratio) * rowWidth + ratio * rowWidth - 18 + thisObj.px;
                            } else {
                                return '35px';
                            }
                        })
                        .classed('linesP', true)
                        .attr('title', thisObj.constantLineValue)
                        .text(function (): string {
                            let displayVal: number;
                            let tempMeasureData: string = thisObj.constantLineValue.toString();
                            const valLen: number = tempMeasureData.split('.')[0].length;
                            displayVal = thisObj.returnDisplayVal(valLen);

                            //Value Formatter Creation
                            const formatter: IValueFormatter = ValueFormatter.create({
                                value: displayVal
                            });

                            tempMeasureData = formatter.format(thisObj.constantLineValue);
                            $(this).attr('title', tempMeasureData);

                            return tempMeasureData;
                        })
                        .style({
                            'font-size': `${this.settings.fontSettings.fontSize}px`,
                            'font-family': this.settings.fontSettings.fontFamily,
                            top: `-${((thisObj.textSize / 4) - 25 + thisObj.textSize)}px`,
                            width: `${thisObj.textWidth}px`
                        });
                }
                thisObj.renderStyles(constantLine + index, divConstantLine, pConstantLine, thisObj.constantLineStyle,
                                     thisObj.constantLineFill, thisObj.constantLineOpacity, thisObj.constantLineValue);
            }
        }

        // tslint:disable-next-line:cyclomatic-complexity no-any
        private renderBinColumnChart(data: any, index: number, categories: any,
                                     chartWidth: number, chartHeight: number, marginLeft: number): void {
            this.noOfColumns = data.length;
            let thisObj: this;
            thisObj = this;
            let indexString: string;
            thisObj.isCategory = true;
            // tslint:disable-next-line:no-any
            let subDiv: any;
            // tslint:disable-next-line:no-any
            let columnChart: any;
            // tslint:disable-next-line:no-any
            let subDiv1: any;
            // tslint:disable-next-line:no-any
            let mainColumnChartContainer: any;
            if (thisObj.colorByColumn !== -1) {
                thisObj.getMaxValue(data);
            } else {
                thisObj.getMaxValue1(data);
            }

            if (thisObj.showXAxisIsOn) {
                const tconstantLine: string = 'constantLine';
            }

            let categoriesLength: number;
            categoriesLength = categories.length;
            let subDivHeight: number;
            let subDivWidth: number;

            const percentSign: string = '%';
            indexString = 'index';
            const divWidth: number = 50;
            const marginToFirstDiv: number = (thisObj.settings.analytics.maxLineDataLabel || thisObj.settings.analytics.minLineDataLabel ||
                thisObj.settings.analytics.avgLineDataLabel || thisObj.settings.analytics.constantLineDataLabel) ?
                thisObj.textWidth + 20 : 30;
            // tslint:disable-next-line:no-any
            //  let innerSubDiv: any;
            const firstDivMargin: number = 30;
            columnChart = thisObj.mainCont.append(`div`)
                .classed('columnChart', true)
                .classed('borderChart', true)
                .style({
                    width: chartWidth + thisObj.px,
                    height: chartHeight + 20 + thisObj.px,
                    'margin-left': marginLeft + thisObj.px
                });

            subDiv1 = columnChart.append('div')
                .style({
                    width: chartWidth + thisObj.px,
                    'min-height': '10px',
                    height: '20px',
                    padding: '0',
                    margin: '0',
                    overflow: 'hidden',
                    'text-align': 'center',
                    'font-weight': 'bold'
                });

            subDiv1.append('label')
                .classed('label', true)
                .text(thisObj.returnLabelText(categories[index]));

            mainColumnChartContainer = columnChart.append('div')
                .classed('mainColumnChartContainer', true)
                .style({
                    position: 'relative',
                    'max-height': '100%',
                    overflow: 'auto',
                    'min-height': '177px',
                    width: chartWidth + thisObj.px
                });

            subDiv = mainColumnChartContainer.append('div')
                .classed('subDiv', true)
                .style({
                    width: (Math.max((data.length * (divWidth + 4)) + marginToFirstDiv, chartWidth - 50)) + 50 + thisObj.px,
                    height: thisObj.returnHeight(categories, chartHeight) - 20 + thisObj.px,
                    padding: '0',
                    margin: '0',
                    overflow: 'auto',
                    'min-height': '150px'
                });

            const $subDiv: JQuery = $('.subDiv');
            subDivHeight = $subDiv.height();
            subDivWidth = $subDiv.width();

            const labelHeight: number = $('.label').height();
            const labelMarginBottom: number =
                parseInt((d3.select('.label').style('margin-bottom').substr(0, d3.select('.label').style('margin-bottom').length - 2)), 10);
            const increment: number = thisObj.previousDataLength;
            // let renderHeight: number;
            // 15 is height of colValueText and 16 is height of Category
            // tslint:disable-next-line:no-any
            //let sum: any;
            const col: string = 'column';
            // tslint:disable-next-line:no-any
            //  let tooltipchart: any;
            // tslint:disable-next-line:no-any
            const tooltipData: any = [];

            // tslint:disable-next-line:no-any
            const innerSubDiv: any = subDiv.selectAll('div')
                .data(data)
                .enter()
                .append('div')
                .classed('mainColumn', true)
                .style('width', divWidth + thisObj.px)
                // tslint:disable-next-line:no-any
                .style('margin-left', function (datum: any, iterator: number): string {
                    if (iterator === 0) {
                        return marginToFirstDiv + thisObj.px;
                    }

                    return '0';
                })
                .style({
                    'margin-right': '4px'
                });
            let renderHeight: number;

            const colValueText: number = 10;
            const categoryHeight: number = 11;
            renderHeight = thisObj.returnRenderHeight(subDivHeight, labelHeight, labelMarginBottom, colValueText, categoryHeight);
            innerSubDiv.append('p')
                .classed('colValueText', true)
                // tslint:disable-next-line:no-any
                .style('margin-top', function (datum: any): string {
                    // tslint:disable-next-line:no-shadowed-variable
                    let styleSum: number = 0;
                    if (thisObj.colorByColumn !== -1) {
                        styleSum = thisObj.getSum(datum);
                    } else {
                        // tslint:disable-next-line:no-string-literal
                        styleSum = datum['values']['value'];
                    }
                    const value: number = styleSum;
                    if (value > 0) {
                        thisObj.isCategory = false;
                    }
                    // tslint:disable-next-line:prefer-template
                    let margin: number = (Math.floor(renderHeight - (styleSum / thisObj.maxValue) * 0.85 * renderHeight));
                    if (styleSum === 0) {
                        // tslint:disable-next-line:prefer-template
                        margin = (Math.floor(renderHeight));
                    }
                    const marginx : number = margin - thisObj.textSize + labelHeight;

                    return (marginx.toString() + thisObj.pixelLiteral);
                })
                .style({
                    'font-size': `${this.settings.fontSettings.fontSize}px`,
                    'font-family': this.settings.fontSettings.fontFamily,
                    'white-space': 'nowrap',
                    overflow: 'hidden',
                    'text-overflow': 'ellipsis',
                    height: `${thisObj.textSize}px`
                })
                // tslint:disable-next-line:no-any
                .text(function (datum: any): string {
                    // tslint:disable-next-line:no-any
                    let displayVal: any;
                    // tslint:disable-next-line:no-shadowed-variable
                    let textSum: number = 0;
                    if (thisObj.colorByColumn !== -1) {
                        textSum = thisObj.getSum(datum);
                    } else {
                        // tslint:disable-next-line:no-string-literal
                        textSum = datum['values']['value'];
                    }
                    let tempMeasureData: string = textSum.toString();
                    const valLen: number = tempMeasureData.split('.')[0].length;
                    displayVal = thisObj.returnDisplayVal(valLen);

                    //Value Formatter Creation
                    const formatter: IValueFormatter = ValueFormatter.create({
                        format: thisObj.dataView.metadata.columns[thisObj.targetColumn].format,
                        value: displayVal
                    });
                    tempMeasureData = formatter.format(textSum);

                    return tempMeasureData;
                })
                // tslint:disable-next-line:no-any
                .attr('title', function (datum: any): string {
                    // tslint:disable-next-line:no-any
                    let attrSum: number;
                    attrSum = 0;
                    if (thisObj.colorByColumn !== -1) {
                        attrSum = thisObj.getSum(datum);
                    } else {
                        // tslint:disable-next-line:no-string-literal
                        attrSum = datum['values']['value'];
                    }
                    // tslint:disable-next-line:no-any
                    let displayVal: any;
                    let tempMeasureData: string = attrSum.toString();
                    const valLen: number = tempMeasureData.split('.')[0].length;
                    displayVal = thisObj.returnDisplayVal(valLen);

                    //Value Formatter Creation
                    const formatter: IValueFormatter = ValueFormatter.create({
                        format: thisObj.dataView.metadata.columns[thisObj.targetColumn].format,
                        value: displayVal
                    });

                    tempMeasureData = formatter.format(attrSum);

                    return tempMeasureData;
                })
                .style('color', thisObj.colors[thisObj.iColumntext][thisObj.jColumntext]);
            innerSubDiv.append('div')
                // tslint:disable-next-line:no-any
                .attr('data-selection', function (datum: any, iterator: number): number { return iterator + increment; })
                .classed('colDiv', true)
                // tslint:disable-next-line:no-any
                .each(function (datum: any): void {
                    // tslint:disable-next-line:typedef
                    const THIS = this;
                    if (thisObj.colorByColumn !== -1) {
                        // tslint:disable-next-line:no-any
                        datum.values[0].forEach(function (d: any, i: number): any {

                            tooltipData.push({ key: d[`key`], value: d[`values`][`value`] });
                            d3.select(THIS)
                                .append('div')
                                .classed('inDiv', true)
                                .style('margin', 0)
                                .style('padding', 0)
                                .style('width', '40px')
                                .style('margin-left', '1px')
                                .each(function (): void {
                                    thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                             (tooltipEvent: TooltipEventArgs<number>) =>
                                            thisObj.getTooltipData(tooltipEvent.data, d, 'bin', '', 0, thisObj),
                                                                             (tooltipEvent: TooltipEventArgs<number>) => null);
                                })
                                .style('height', function (): string {
                                    thisObj.globalSelections.push({
                                        data: datum,
                                        binIndex: index,
                                        category: categories[index]
                                    });
                                    // tslint:disable-next-line:no-string-literal
                                    const styleSum: number = d['values']['value'];

                                    return styleSum <= 0 ? `0` :
                                        `${((styleSum / thisObj.maxValue) * 0.85 * renderHeight)}px`;
                                })
                                // tslint:disable-next-line:no-any
                                .style('background-color', function (): any {
                                    // tslint:disable-next-line:no-any
                                    return thisObj.barcolor.filter(function (v: any): any {
                                        return v.key.toString() === d[`key`].toString();
                                    })[0].value;
                                });
                        });
                    } else {
                        d3.select(this)
                            .style('width', '40px')
                            .style('margin-left', '1px')
                            // tslint:disable-next-line:no-any
                            .style('height', function (tdatum: any): string {
                                // tslint:disable-next-line:no-any
                                let elseSum: any;
                                // tslint:disable-next-line:no-string-literal
                                elseSum = tdatum['values']['value'];
                                if (thisObj.binColumn === -1) {

                                    return elseSum <= 0 ? `0` :
                                        `${Math.floor((elseSum / thisObj.maxValue * (
                                            0.85 * renderHeight)))}px`;
                                }

                                return elseSum <= 0 ? `0` :
                                    `${Math.floor((elseSum / thisObj.maxValue * (
                                        0.85 * renderHeight)))}px`;
                            })
                            // tslint:disable-next-line:no-any
                            .each(function (tdatum: any): void {
                                thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                         (tooltipEvent: TooltipEventArgs<number>) =>
                                        thisObj.getTooltipData(tooltipEvent.data, tdatum, 'bin', '', 0, thisObj),
                                                                         (tooltipEvent: TooltipEventArgs<number>) => null);
                            })
                            .style('background-color', thisObj.colors[thisObj.iColumn][thisObj.jColumn]);
                    }
                });

            // append x-axis labels (Category text)
            innerSubDiv.append('p')
                .classed('category', true)
                // tslint:disable-next-line:no-any
                .text(function (datum: any): string { return datum[`key`].toString(); })
                .style({
                    'font-size': `${this.settings.fontSettings.fontSize}px`,
                    'font-family': this.settings.fontSettings.fontFamily,
                    height: `${thisObj.textSize}px`
                })
                // tslint:disable-next-line:no-any
                .attr('title', function (datum: any): string { return datum[`key`].toString(); });

            const border: number = 2;
            const colValueTextHeight: number = $('.colValueText').height();
            const minLine: string = 'columnMinLine';
            const maxLine: string = 'columnMaxLine';
            const constantLine: string = 'columnConstantLine';
            const avgLine: string = 'columnAvgLine';
            const divMaxLine: string = 'max';
            const pMaxLine: string = 'pMax';
            const divMinLine: string = 'min';
            const pMinLine: string = 'pMin';
            const divAvgLine: string = 'avg';
            const pAvgLine: string = 'pAvg';
            const divConstantLine: string = 'constant';
            const pConstantLine: string = 'pConstant';
            const max: number = thisObj.maxValue;
            const min: number = thisObj.minValue;
            const average: number = thisObj.averageValue;
            const chartWdth: number = thisObj.returnChartWidth(chartWidth);

            const updatedLeftMargin: () => string = function (): string {
                return (thisObj.maxLineDataLabel === true || thisObj.minLineDataLabel === true
                    || thisObj.avgLineDataLabel === true || thisObj.constantLineDataLabel === true) ? '51px' : '21px';
            };

            if (thisObj.maxLineIsOn) {
                if (thisObj.maxLineDataLabel) {
                    subDiv.append('p')
                        .classed(pMaxLine, true)
                        // subtract 25 to adjust label above the line
                        .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                            labelHeight + labelMarginBottom - thisObj.textSize + thisObj.px
                            : labelHeight + labelMarginBottom +
                            (Math.floor(renderHeight - 0.85 * renderHeight)) - thisObj.textSize + thisObj.px)
                        .classed('linesP', true)
                        .style({
                            'font-size': `${this.settings.fontSettings.fontSize}px`,
                            'font-family': this.settings.fontSettings.fontFamily,
                            width: `${thisObj.textWidth}px`,
                            margin: '0'
                        })
                        .text(function (): string {
                            let displayVal: number;
                            let tempMeasureData: string = thisObj.maxValue.toString();
                            const valLen: number = tempMeasureData.split('.')[0].length;
                            displayVal = thisObj.returnDisplayVal(valLen);

                            //Value Formatter Creation
                            const formatter: IValueFormatter = ValueFormatter.create({
                                value: displayVal
                            });

                            tempMeasureData = formatter.format(thisObj.maxValue);
                            $(this).attr('title', tempMeasureData);

                            return tempMeasureData;
                        });
                }

                // tslint:disable-next-line:no-any
                subDiv.append('div')
                    .classed(divMaxLine, true)
                    .classed(maxLine + index, true)
                    .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                        labelHeight + labelMarginBottom + thisObj.px
                        : labelHeight + labelMarginBottom +
                        (Math.floor(renderHeight - 0.85 * renderHeight)) + thisObj.px)
                    .style('margin-left', updatedLeftMargin())
                    .style('width', function (): string {
                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                thisObj.getTooltipDataAnalyticsLine(tooltipEvent.data, '', 'line', 'Max', max, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);

                        return (((thisObj.noOfColumns * 53) + thisObj.noOfColumns + 7) + thisObj.px);
                    });

                thisObj.renderStyles(maxLine + index, divMaxLine, pMaxLine, thisObj.maxLineStyle,
                                     thisObj.maxLineFill, thisObj.maxLineOpacity, thisObj.maxValue);
            }
            if (thisObj.minLineIsOn) {
                if (thisObj.minLineDataLabel) {
                    subDiv.append('p')
                        .classed(pMinLine, true)
                        .style({
                            'font-size': `${this.settings.fontSettings.fontSize}px`,
                            'font-family': this.settings.fontSettings.fontFamily,
                            width: `${thisObj.textWidth}px`,
                            margin: 0
                        })
                        .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                            labelHeight + labelMarginBottom + 25 - thisObj.textSize + thisObj.px
                            : labelHeight + labelMarginBottom + 25 - thisObj.textSize +
                            (Math.floor(renderHeight - (thisObj.minValue / thisObj.maxValue) * 0.85 * renderHeight)) - 25 + thisObj.px)
                        .classed('linesP', true)
                        .text(function (): string {
                            let displayVal: number;
                            let tempMeasureData: string = thisObj.minValue.toString();
                            const valLen: number = tempMeasureData.split('.')[0].length;
                            displayVal = thisObj.returnDisplayVal(valLen);

                            //Value Formatter Creation
                            const formatter: IValueFormatter = ValueFormatter.create({
                                value: displayVal
                            });

                            tempMeasureData = formatter.format(thisObj.minValue);
                            $(this).attr('title', tempMeasureData);

                            return tempMeasureData;
                        });
                }
                subDiv.append('div')
                    .classed(divMinLine, true)
                    .classed(minLine + index, true)
                    .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                        labelHeight + labelMarginBottom + thisObj.px
                        : labelHeight + labelMarginBottom +
                        (Math.floor(renderHeight - (thisObj.minValue / thisObj.maxValue) * 0.85 * renderHeight)) + thisObj.px)
                    .style('margin-left', updatedLeftMargin())
                    .style('width', function (): string {
                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                thisObj.getTooltipDataAnalyticsLine(tooltipEvent.data, '', 'line', 'Min', min, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);

                        return (((thisObj.noOfColumns * 53) + thisObj.noOfColumns + 7) + thisObj.px);
                    });
                thisObj.renderStyles(minLine + index, divMinLine, pMinLine, thisObj.minLineStyle,
                                     thisObj.minLineFill, thisObj.minLineOpacity, thisObj.minValue);
            }
            if (thisObj.avgLineIsOn) {
                if (thisObj.avgLineDataLabel) {
                    subDiv.append('p')
                        .classed(pAvgLine, true)
                        .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                            labelHeight + labelMarginBottom - thisObj.textSize + thisObj.px
                            : labelHeight + labelMarginBottom +
                            (Math.floor(renderHeight - (thisObj.averageValue / thisObj.maxValue) * 0.85 * renderHeight))
                            - thisObj.textSize + thisObj.px)
                        .classed('linesP', true)
                        .style({
                            'font-size': `${this.settings.fontSettings.fontSize}px`,
                            'font-family': this.settings.fontSettings.fontFamily,
                            width: `${thisObj.textWidth}px`,
                            margin: 0
                        })
                        .text(function (): string {
                            let displayVal: number;
                            let tempMeasureData: string = thisObj.averageValue.toString();
                            const valLen: number = tempMeasureData.split('.')[0].length;
                            displayVal = thisObj.returnDisplayVal(valLen);

                            //Value Formatter Creation
                            const formatter: IValueFormatter = ValueFormatter.create({
                                value: displayVal
                            });

                            tempMeasureData = formatter.format(thisObj.averageValue);
                            $(this).attr('title', tempMeasureData);

                            return tempMeasureData;
                        });
                }
                subDiv.append('div')
                    .classed(avgLine + index, true)
                    .classed(divAvgLine, true)
                    .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                        labelHeight + labelMarginBottom + thisObj.px
                        : labelHeight + labelMarginBottom +
                        (Math.floor(renderHeight - (thisObj.averageValue / thisObj.maxValue) * 0.85 * renderHeight)) + thisObj.px)
                    .style('margin-left', updatedLeftMargin())
                    .style('width', function (): string {
                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                thisObj.getTooltipDataAnalyticsLine(tooltipEvent.data, '', 'line', 'Average', average, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);

                        return (((thisObj.noOfColumns * 53) + thisObj.noOfColumns + 7) + thisObj.px);
                    });
                thisObj.renderStyles(avgLine + index, divAvgLine, pAvgLine, thisObj.avgLineStyle,
                                     thisObj.avgLineFill, thisObj.avgLineOpacity, thisObj.averageValue);

            }
            if (thisObj.constantLineIsOn) {
                if (thisObj.constantLineDataLabel) {
                    subDiv.append('p')
                        .classed(pConstantLine, true)
                        .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                            labelHeight + labelMarginBottom - thisObj.textSize + thisObj.px
                            : labelHeight + labelMarginBottom +
                            (Math.floor(renderHeight - (thisObj.constantLineValue / thisObj.maxValue) * 0.85 * renderHeight))
                            - thisObj.textSize + thisObj.px)
                        .classed('linesP', true)
                        .style({
                            'font-size': `${this.settings.fontSettings.fontSize}px`,
                            'font-family': this.settings.fontSettings.fontFamily,
                            width: `${thisObj.textWidth}px`,
                            margin: 0
                        })
                        .text(function (): string {
                            let displayVal: number;
                            let tempMeasureData: string = thisObj.constantLineValue.toString();
                            const valLen: number = tempMeasureData.split('.')[0].length;
                            displayVal = thisObj.returnDisplayVal(valLen);

                            //Value Formatter Creation
                            const formatter: IValueFormatter = ValueFormatter.create({
                                value: displayVal
                            });

                            tempMeasureData = formatter.format(thisObj.constantLineValue);
                            $(this).attr('title', tempMeasureData);

                            return tempMeasureData;
                        });
                }
                subDiv.append('div')
                    .classed(constantLine + index, true)
                    .classed(divConstantLine, true)
                    .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                        labelHeight + labelMarginBottom + thisObj.px
                        : labelHeight + labelMarginBottom +
                        (Math.floor(renderHeight - (thisObj.constantLineValue / thisObj.maxValue) * 0.85 * renderHeight))
                        + thisObj.px)
                    .style('margin-left', updatedLeftMargin())
                    .style('width', ((this.noOfColumns * 53) + this.noOfColumns + 7) + thisObj.px)
                    .each(function (): void {
                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                thisObj.getTooltipDataAnalyticsLine(tooltipEvent.data, '', 'line', 'Max', thisObj.maxValue, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);
                    });
                thisObj.renderStyles(constantLine + index, divConstantLine, pConstantLine, thisObj.constantLineStyle,
                                     thisObj.constantLineFill, thisObj.constantLineOpacity, thisObj.constantLineValue);
            }

            if (thisObj.showXAxisIsOn) {
                subDiv.append('div')
                    .classed(divConstantLine, true)
                    .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                        labelHeight + labelMarginBottom + thisObj.px
                        : labelHeight + labelMarginBottom +
                        (Math.floor(renderHeight - (thisObj.constantLineValue / thisObj.maxValue) * 0.85 * renderHeight))
                        + thisObj.px)

                    //.style('margin-left', updatedLeftMargin())
                    .style('margin-left', `${thisObj.textWidth}px`)
                    .style('width', ((this.noOfColumns * 53) + this.noOfColumns + 7) + thisObj.px);
                thisObj.renderStyles(constantLine + index, divConstantLine, pConstantLine, thisObj.constantLineStyle,
                                     thisObj.constantLineFill, thisObj.constantLineOpacity, thisObj.constantLineValue);
            }

            // On Load Animation
            if (thisObj.renderedTime === 1) {
                d3.selectAll('.columnChart .myDiv .mainColumn')
                    .style('transform-origin', '0% 0%')
                    .style('transform', 'scale(0)')
                    .transition()
                    .duration(2000)
                    .style('transform', 'scale(1)');
            }
            subDiv.style('transform', 'scale(0)')
                .transition().duration(500)
                .style('transform', 'scale(1)');
        }

        // tslint:disable-next-line:cyclomatic-complexity no-any
        private renderColumnChart(data: any, index: number, categories: any,
                                  chartWidth: number, chartHeight: number, marginLeft: number): void {
            this.noOfColumns = data.length;

            const thisObj: this = this;
            const indexString: string = 'index';
            thisObj.isCategory = true;
            // tslint:disable-next-line:no-any
            let subDiv: any;
            // tslint:disable-next-line:no-any
            let columnChart: any;
            // tslint:disable-next-line:no-any
            let subDiv1: any;
            // tslint:disable-next-line:no-any
            let mainColumnChartContainer: any;
            if (thisObj.colorByColumn !== -1) {
                thisObj.getMaxValue(data);
            } else {
                thisObj.getMaxValue1(data);
            }
            const categoriesLength: number = categories.length;
            let subDivHeight: number;
            let subDivWidth: number;
            const percentSign: string = '%';
            const divWidth: number = 50;
            const marginToFirstDiv: number = (thisObj.settings.analytics.maxLineDataLabel || thisObj.settings.analytics.minLineDataLabel ||
                thisObj.settings.analytics.avgLineDataLabel || thisObj.settings.analytics.constantLineDataLabel) ?
                thisObj.textWidth + 20 : 30;
            const firstDivMargin: number = 30;
            if (thisObj.binColumn === -1) {
                columnChart = thisObj.mainCont.append(`div`)
                    .classed('columnChart', true);
                subDiv = columnChart.append('div')
                    .classed('subDiv', true)
                    .style({
                        /* Updated formula to avoid extra space */
                        width: (Math.max((data.length * (divWidth + 4)) + marginToFirstDiv, chartWidth - 50)) + 50 + thisObj.px,
                        height: thisObj.mainContHeight - $('.legend1').height() - + thisObj.px,
                        padding: '0',
                        margin: '0',
                        'min-height': '150px'
                    });

            } else {
                //added one sub div for label
                columnChart = thisObj.mainCont.append(`div`)
                    .classed('columnChart', true)
                    .classed('borderChart', true)
                    .style({
                        width: chartWidth + thisObj.px,
                        height: chartHeight + 20 + thisObj.px,
                        'margin-left': marginLeft + thisObj.px
                    });

                subDiv1 = columnChart.append('div')
                    .style({
                        width: chartWidth + thisObj.px,
                        'min-height': '10px',
                        height: '20px',
                        padding: '0',
                        margin: '0',
                        overflow: 'hidden',
                        'text-align': 'center',
                        'font-weight': 'bold'
                    });

                subDiv1.append('label')
                    .classed('label', true)
                    .text(thisObj.returnLabelText(categories[index]));

                mainColumnChartContainer = columnChart.append('div')
                    .classed('mainColumnChartContainer', true)
                    .style({
                        position: 'relative',
                        'max-height': '100%',
                        overflow: 'auto',
                        'min-height': '177px',
                        width: chartWidth + thisObj.px
                    });

                subDiv = mainColumnChartContainer.append('div')
                    .classed('subDiv', true)
                    .style({
                        width: (Math.max((data.length * (divWidth + 4)) + marginToFirstDiv, chartWidth - 50)) + 50 + thisObj.px,
                        height: thisObj.returnHeight(categories, chartHeight) - 20 + thisObj.px,
                        padding: '0',
                        margin: '0',
                        overflow: 'auto',
                        'min-height': '150px'
                    });
            }
            const $subDiv: JQuery = $('.subDiv');
            subDivHeight = $subDiv.height();
            subDivWidth = $subDiv.width();

            const labelHeight: number = thisObj.binColumn === -1 ? 17 : $('.label').height();
            const labelMarginBottom: number = thisObj.binColumn === -1 ? 0 :
                parseInt((d3.select('.label').style('margin-bottom').substr(0, d3.select('.label').style('margin-bottom').length - 2)), 10);

            const increment: number = thisObj.previousDataLength;
            // 15 is height of colValueText and 16 is height of Category

            // tslint:disable-next-line:no-any
            const col: string = 'column';
            // tslint:disable-next-line:no-any
            //   const tooltipchart: any;
            // tslint:disable-next-line:no-any
            const tooltipData: any = [];
            // innerSubDiv

            // tslint:disable-next-line:no-any
            const innerSubDiv: any = subDiv.selectAll('div')
                .data(data)
                .enter()
                .append('div')
                .classed('mainColumn', true)
                .style('width', divWidth + thisObj.px)
                // tslint:disable-next-line:no-any
                .style('margin-left', function (datum: any, iterator: number): string {
                    if (iterator === 0) {
                        return marginToFirstDiv + thisObj.px;
                    }

                    return '0';
                })
                .style({
                    'margin-right': '4px'
                });
            let renderHeight: number;
            // 15 is height of colValueText and 16 is height of Category
            const colValueText: number = 15;
            const categoryHeight: number = 16;
            renderHeight = thisObj.returnRenderHeight(subDivHeight, labelHeight, labelMarginBottom, colValueText, categoryHeight);
            innerSubDiv.append('p')
                .classed('colValueText', true)
                // tslint:disable-next-line:no-any
                .style('margin-top', function (datum: any): string {
                    let styleSum: number = 0;
                    if (thisObj.colorByColumn !== -1) {
                        styleSum = thisObj.getSum(datum);
                    } else {
                        // tslint:disable-next-line:no-string-literal
                        styleSum = datum['values']['value'];
                    }
                    const value: number = styleSum;
                    if (value > 0) {
                        thisObj.isCategory = false;
                    }
                    // tslint:disable-next-line:prefer-template
                    let margin: number = (Math.floor(renderHeight - (styleSum / thisObj.maxValue) * 0.85 * renderHeight));
                    if (styleSum === 0) {
                        // tslint:disable-next-line:prefer-template
                        margin = (Math.floor(renderHeight));
                    }
                    const marginx : number = margin - thisObj.textSize + labelHeight;

                    return (marginx.toString() + thisObj.pixelLiteral);
                })
                .style({
                    'font-size': `${this.settings.fontSettings.fontSize}px`,
                    'font-family': this.settings.fontSettings.fontFamily,
                    'white-space': 'nowrap',
                    overflow: 'hidden',
                    'text-overflow': 'ellipsis',
                    height: `${thisObj.textSize}px`
                })
                // tslint:disable-next-line:no-any
                .text(function (datum: any): string {
                    // tslint:disable-next-line:no-any
                    let displayVal: any;
                    let textSum: number = 0;
                    if (thisObj.colorByColumn !== -1) {
                        textSum = thisObj.getSum(datum);
                    } else {
                        // tslint:disable-next-line:no-string-literal
                        textSum = datum['values'].value;
                    }
                    let tempMeasureData: string = textSum.toString();
                    const valLen: number = tempMeasureData.split('.')[0].length;
                    displayVal = thisObj.returnDisplayVal(valLen);

                    //Value Formatter Creation
                    const formatter: IValueFormatter = ValueFormatter.create({
                        format: thisObj.dataView.metadata.columns[thisObj.targetColumn].format,
                        value: displayVal
                    });
                    tempMeasureData = formatter.format(textSum);

                    return tempMeasureData;
                })
                // tslint:disable-next-line:no-any
                .attr('title', function (datum: any): string {
                    // tslint:disable-next-line:no-any
                    let attrSum: number = 0;
                    if (thisObj.colorByColumn !== -1) {
                        attrSum = thisObj.getSum(datum);
                    } else {
                        // tslint:disable-next-line:no-string-literal
                        attrSum = datum['values'].value;
                    }
                    // tslint:disable-next-line:no-any
                    let displayVal: any;
                    let tempMeasureData: string = attrSum.toString();
                    const valLen: number = tempMeasureData.split('.')[0].length;
                    displayVal = thisObj.returnDisplayVal(valLen);

                    //Value Formatter Creation
                    const formatter: IValueFormatter = ValueFormatter.create({
                        format: thisObj.dataView.metadata.columns[thisObj.targetColumn].format,
                        value: displayVal
                    });

                    tempMeasureData = formatter.format(attrSum);

                    return tempMeasureData;
                })
                .style('color', thisObj.colors[thisObj.iColumntext][thisObj.jColumntext]);
            innerSubDiv.append('div')
                // tslint:disable-next-line:no-any
                .attr('data-selection', function (datum: any, iterator: number): number { return iterator + increment; })
                .classed('colDiv', true)
                // tslint:disable-next-line:no-any
                .each(function (datum: any): void {
                    // tslint:disable-next-line:no-any
                    const THIS: any = this;
                    if (thisObj.colorByColumn !== -1) {
                        let innerSubDivsum: number;
                        innerSubDivsum = thisObj.getSum(datum);
                        // tslint:disable-next-line:no-any
                        datum.values[0].forEach(function (d: any, i: number): any {
                            tooltipData.push({ key: d[`key`], value: d[`values`][`value`] });
                            d3.select(THIS)
                                .append('div')
                                .classed('inDiv', true)
                                .style('margin', 0)
                                .style('padding', 0)
                                .style('width', '40px')
                                .style('margin-left', '1px')
                                // tslint:disable-next-line:no-any
                                .each(function (): void {
                                    thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                             (tooltipEvent: TooltipEventArgs<number>) =>
                                            thisObj.getTooltipData(tooltipEvent.data, d, 'bin', '', 0, thisObj),
                                                                             (tooltipEvent: TooltipEventArgs<number>) => null);
                                })
                                .style('height', function (): string {
                                    // tslint:disable-next-line:no-any
                                    let styleSum: any;
                                    // tslint:disable-next-line:no-string-literal
                                    styleSum = d['values'].value;

                                    return styleSum <= 0 ? `0` :
                                        `${((styleSum / thisObj.maxValue) * 0.85 * renderHeight)}px`;
                                })
                                // tslint:disable-next-line:no-any
                                .style('background-color', function (): any {
                                    // tslint:disable-next-line:no-any
                                    return thisObj.barcolor.filter(function (v: any): any {
                                        return v.key.toString() === d[`key`].toString();
                                    })[0].value;
                                });
                        });
                    } else {
                        d3.select(this)
                            .style('width', '40px')
                            .style('margin-left', '1px')
                            // tslint:disable-next-line:no-any
                            .style('height', function (tdatum: any): string {
                                let styleSum: number;
                                // tslint:disable-next-line:no-string-literal
                                styleSum = tdatum['values'].value;
                                if (thisObj.binColumn === -1) {
                                    return styleSum <= 0 ? `0` :
                                        `${Math.floor((styleSum / thisObj.maxValue * (
                                            0.85 * renderHeight)))}px`;
                                }

                                return styleSum <= 0 ? `0` :
                                    `${Math.floor((styleSum / thisObj.maxValue * (
                                        0.85 * renderHeight)))}px`;
                            })
                            // tslint:disable-next-line:no-any
                            .each(function (): void {
                                thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                         (tooltipEvent: TooltipEventArgs<number>) =>
                                        thisObj.getTooltipData(tooltipEvent.data, data, 'bin', '', 0, thisObj),
                                                                         (tooltipEvent: TooltipEventArgs<number>) => null);
                            })
                            .style('background-color', thisObj.colors[thisObj.iColumn][thisObj.jColumn]);
                    }
                });

            // append x-axis labels (Category text)
            innerSubDiv.append('p')
                .classed('category', true)
                // tslint:disable-next-line:no-any
                .text(function (datum: any): string { return datum[`key`].toString(); })
                .style({
                    'font-size': `${this.settings.fontSettings.fontSize}px`,
                    'font-family': this.settings.fontSettings.fontFamily,
                    height: `${thisObj.textSize}px`
                })
                // tslint:disable-next-line:no-any
                .attr('title', function (datum: any): string { return datum[`key`].toString(); });

            const border: number = 2;
            const colValueTextHeight: number = $('.colValueText').height();
            const minLine: string = 'columnMinLine';
            const maxLine: string = 'columnMaxLine';
            const constantLine: string = 'columnConstantLine';
            const avgLine: string = 'columnAvgLine';
            const divMaxLine: string = 'max';
            const pMaxLine: string = 'pMax';
            const divMinLine: string = 'min';
            const pMinLine: string = 'pMin';
            const divAvgLine: string = 'avg';
            const pAvgLine: string = 'pAvg';
            const divConstantLine: string = 'constant';
            const pConstantLine: string = 'pConstant';
            const max: number = thisObj.maxValue;
            const min: number = thisObj.minValue;
            const average: number = thisObj.averageValue;
            const chartWdth: number = thisObj.returnChartWidth(chartWidth);
            const updatedLeftMargin: () => string = function (): string {
                return (thisObj.maxLineDataLabel === true || thisObj.minLineDataLabel === true
                    || thisObj.avgLineDataLabel === true || thisObj.constantLineDataLabel === true) ? `${thisObj.textWidth + 5}px` : '21px';
            };
            if (thisObj.maxLineIsOn) {
                if (thisObj.maxLineDataLabel) {

                    subDiv.append('p')
                        .classed(pMaxLine, true)
                        // subtract 25 to adjust label above the line
                        .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                            labelHeight + labelMarginBottom - thisObj.textSize + thisObj.px
                            : labelHeight + labelMarginBottom +
                            (Math.floor(renderHeight - 0.85 * renderHeight)) - thisObj.textSize + thisObj.px)
                        .classed('linesP', true)
                        .style({
                            'font-size': `${this.settings.fontSettings.fontSize}px`,
                            'font-family': this.settings.fontSettings.fontFamily,
                            width: `${thisObj.textWidth}px`,
                            margin: '0'
                        })
                        .text(function (): string {
                            let displayVal: number;
                            let tempMeasureData: string = thisObj.maxValue.toString();
                            const valLen: number = tempMeasureData.split('.')[0].length;
                            displayVal = thisObj.returnDisplayVal(valLen);

                            //Value Formatter Creation
                            const formatter: IValueFormatter = ValueFormatter.create({
                                value: displayVal
                            });

                            tempMeasureData = formatter.format(thisObj.maxValue);
                            $(this).attr('title', tempMeasureData);

                            return tempMeasureData;
                        });
                }
                subDiv.append('div')
                    .classed(divMaxLine, true)
                    .classed(maxLine + index, true)
                    .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                        labelHeight + 16 + thisObj.px
                        : labelHeight + labelMarginBottom  +
                        (Math.floor(renderHeight - 0.85 * renderHeight)) + thisObj.px)
                    .style('margin-left', updatedLeftMargin())
                    .style('width', function (): string {
                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                thisObj.getTooltipDataAnalyticsLine(tooltipEvent.data, '', 'line', 'Max', max, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);

                        return (((thisObj.noOfColumns * 53) + thisObj.noOfColumns + 7) + thisObj.px);
                    });
                thisObj.renderStyles(maxLine + index, divMaxLine, pMaxLine, thisObj.maxLineStyle,
                                     thisObj.maxLineFill, thisObj.maxLineOpacity, thisObj.maxValue);
            }
            if (thisObj.minLineIsOn) {
                if (thisObj.minLineDataLabel) {
                    subDiv.append('p')
                        .classed(pMinLine, true)
                        .style({
                            'font-size': `${this.settings.fontSettings.fontSize}px`,
                            'font-family': this.settings.fontSettings.fontFamily,
                            width: `${thisObj.textWidth}px`,
                            margin: 0
                        })
                        .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                            labelHeight + labelMarginBottom + 25 - thisObj.textSize + thisObj.px
                            : labelHeight + labelMarginBottom + 25 - thisObj.textSize +
                            (Math.floor(renderHeight - (thisObj.minValue / thisObj.maxValue) * 0.85 * renderHeight)) - 25 + thisObj.px)
                        .classed('linesP', true)
                        .text(function (): string {
                            let displayVal: number;
                            let tempMeasureData: string = thisObj.minValue.toString();
                            const valLen: number = tempMeasureData.split('.')[0].length;
                            displayVal = thisObj.returnDisplayVal(valLen);

                            //Value Formatter Creation
                            const formatter: IValueFormatter = ValueFormatter.create({
                                value: displayVal
                            });

                            tempMeasureData = formatter.format(thisObj.minValue);
                            $(this).attr('title', tempMeasureData);

                            return tempMeasureData;
                        });
                }
                subDiv.append('div')
                    .classed(divMinLine, true)
                    .classed(minLine + index, true)
                    // .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                    //     labelHeight + labelMarginBottom + thisObj.px
                    //     : labelHeight + labelMarginBottom +
                    //     (Math.floor(renderHeight - (thisObj.minValue / thisObj.maxValue) * 0.85 * renderHeight)) + thisObj.px)
                    .style('margin-top', function (): string {
                        let styleSum: number;
                        styleSum = (thisObj.isCategory ? (Math.floor(renderHeight)) :
                            (Math.floor(renderHeight - (thisObj.minValue / thisObj.maxValue) * 0.85 * renderHeight)));
                        let totalSum: number;
                        totalSum = labelHeight + labelMarginBottom + styleSum;

                        // tslint:disable-next-line:prefer-template
                        return totalSum + 'px';
                    })
                    .style('margin-left', updatedLeftMargin())
                    .style('width', function (): string {
                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                thisObj.getTooltipDataAnalyticsLine(tooltipEvent.data, '', 'line', 'Min', min, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);

                        return (((thisObj.noOfColumns * 53) + thisObj.noOfColumns + 7) + thisObj.px);
                    });
                thisObj.renderStyles(minLine + index, divMinLine, pMinLine, thisObj.minLineStyle,
                                     thisObj.minLineFill, thisObj.minLineOpacity, thisObj.minValue);
            }
            if (thisObj.avgLineIsOn) {
                if (thisObj.avgLineDataLabel) {
                    subDiv.append('p')
                        .classed(pAvgLine, true)
                        .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                            labelHeight + labelMarginBottom - thisObj.textSize + thisObj.px
                            : labelHeight + labelMarginBottom +
                            (Math.floor(renderHeight - (thisObj.averageValue / thisObj.maxValue) * 0.85 * renderHeight))
                            - thisObj.textSize + thisObj.px)
                        .classed('linesP', true)
                        .style({
                            'font-size': `${this.settings.fontSettings.fontSize}px`,
                            'font-family': this.settings.fontSettings.fontFamily,
                            width: `${thisObj.textWidth}px`,
                            margin: 0
                        })
                        .text(function (): string {
                            let displayVal: number;
                            let tempMeasureData: string = thisObj.averageValue.toString();
                            const valLen: number = tempMeasureData.split('.')[0].length;
                            displayVal = thisObj.returnDisplayVal(valLen);

                            //Value Formatter Creation
                            const formatter: IValueFormatter = ValueFormatter.create({
                                value: displayVal
                            });

                            tempMeasureData = formatter.format(thisObj.averageValue);
                            $(this).attr('title', tempMeasureData);

                            return tempMeasureData;
                        });
                }
                subDiv.append('div')
                    .classed(avgLine + index, true)
                    .classed(divAvgLine, true)
                    .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                        labelHeight + labelMarginBottom + thisObj.px
                        : labelHeight + labelMarginBottom +
                        (Math.floor(renderHeight - (thisObj.averageValue / thisObj.maxValue) * 0.85 * renderHeight)) + thisObj.px)
                    .style('margin-left', updatedLeftMargin())

                    .style('width', function (): string {
                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                thisObj.getTooltipDataAnalyticsLine(tooltipEvent.data, '', 'line', 'Average', average, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);

                        return (((thisObj.noOfColumns * 53) + thisObj.noOfColumns + 7) + thisObj.px);

                        //.style('margin-left', "21px")
                        //.style('width', ((this.noOfColumns * 53)  + this.noOfColumns+7 )+ thisObj.px)

                    });
                thisObj.renderStyles(avgLine + index, divAvgLine, pAvgLine, thisObj.avgLineStyle,
                                     thisObj.avgLineFill, thisObj.avgLineOpacity, thisObj.averageValue);

            }
            if (thisObj.constantLineIsOn) {

                if (thisObj.constantLineDataLabel) {
                    subDiv.append('p')
                        .classed(pConstantLine, true)
                        .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                            labelHeight + labelMarginBottom - thisObj.textSize + thisObj.px
                            : labelHeight + labelMarginBottom +
                            (Math.floor(renderHeight - (thisObj.constantLineValue / thisObj.maxValue) * 0.85 * renderHeight))
                            - thisObj.textSize + thisObj.px)
                        .style('margin-top', '533px')
                        .classed('linesP', true)
                        .style({
                            'font-size': `${this.settings.fontSettings.fontSize}px`,
                            'font-family': this.settings.fontSettings.fontFamily,
                            width: `${thisObj.textWidth}px`,
                            margin: 0
                        })
                        .text(function (): string {
                            let displayVal: number;
                            let tempMeasureData: string = thisObj.constantLineValue.toString();
                            const valLen: number = tempMeasureData.split('.')[0].length;
                            displayVal = thisObj.returnDisplayVal(valLen);

                            //Value Formatter Creation
                            const formatter: IValueFormatter = ValueFormatter.create({
                                value: displayVal
                            });

                            tempMeasureData = formatter.format(thisObj.constantLineValue);
                            $(this).attr('title', tempMeasureData);

                            return tempMeasureData;
                        });
                }
                subDiv.append('div')
                    .classed(constantLine + index, true)
                    .classed(divConstantLine, true)
                    // tslint:disable-next-line:no-any
                    .style('margin-top', function (): any {
                        let subDivSum: number;
                        subDivSum = (thisObj.isCategory ? (Math.floor(renderHeight)) :
                            (Math.floor(renderHeight - (thisObj.constantLineValue / thisObj.maxValue) * 0.85 * renderHeight)));
                        let totalSum: number;
                        totalSum = labelHeight + labelMarginBottom + subDivSum;

                        // tslint:disable-next-line:prefer-template
                        return totalSum + 'px';
                    })
                    .style('margin-left', updatedLeftMargin())
                    //.style('width', (Math.max((data.length * (divWidth + 4)) + marginToFirstDiv - 40, chartWidth))  + thisObj.px)
                    .style('width', ((this.noOfColumns * 53) + this.noOfColumns + 7) + thisObj.px)
                    //.style('height', ((this.noOfBars * 15) + this.noOfBars - 2) + thisObj.px)
                    .each(function (): void {
                        thisObj.tooltipServiceWrapper.addTooltip(d3.select(this),
                                                                 (tooltipEvent: TooltipEventArgs<number>) =>
                                // tslint:disable-next-line:max-line-length
                                thisObj.getTooltipDataAnalyticsLine(tooltipEvent.data, '', 'line', 'constant', thisObj.constantLineValue, thisObj),
                                                                 (tooltipEvent: TooltipEventArgs<number>) => null);
                    });
                thisObj.renderStyles(constantLine + index, divConstantLine, pConstantLine, thisObj.constantLineStyle,
                                     thisObj.constantLineFill, thisObj.constantLineOpacity, thisObj.constantLineValue);
            }

            if (thisObj.showXAxisIsOn) {
                if (thisObj.binColumn === -1) {
                    subDiv.append('div')
                        .classed(divConstantLine, true)
                        // tslint:disable-next-line:no-any
                        .style('margin-top', function (): any {
                            let styleSum: number;
                            styleSum = (thisObj.isCategory ? (Math.floor(renderHeight)) :
                                (Math.floor(renderHeight - (thisObj.constantLineValue / thisObj.maxValue) * 0.85 * renderHeight)));
                            let totalSum: number;
                            totalSum = labelHeight + labelMarginBottom + styleSum + 0;

                            // tslint:disable-next-line:prefer-template
                            return totalSum + 'px';
                        })
                        .style('margin-left', updatedLeftMargin())
                        .style('width', ((this.noOfColumns * 53) + this.noOfColumns + 7) + thisObj.px);
                } else {
                    subDiv.append('div')
                        .classed(divConstantLine, true)
                        .style('top', thisObj.isCategory ? Math.floor(renderHeight) +
                            labelHeight + labelMarginBottom + thisObj.px
                            : labelHeight + labelMarginBottom +
                            (Math.floor(renderHeight - (thisObj.constantLineValue / thisObj.maxValue) * 0.85 * renderHeight))
                            + thisObj.px)

                        .style('margin-left', updatedLeftMargin())
                        .style('width', ((this.noOfColumns * 53) + this.noOfColumns + 7) + thisObj.px);
                }
                thisObj.renderStyles(constantLine + index, divConstantLine, pConstantLine, thisObj.constantLineStyle,
                                     thisObj.constantLineFill, thisObj.constantLineOpacity, thisObj.constantLineValue);
            }

            // On Load Animation
            if (thisObj.renderedTime === 1) {
                d3.selectAll('.columnChart .myDiv .mainColumn')
                    .style('transform-origin', '0% 0%')
                    .style('transform', 'scale(0)')
                    .transition()
                    .duration(2000)
                    .style('transform', 'scale(1)');
            }
            subDiv.style('transform', 'scale(0)')
                .transition().duration(500)
                .style('transform', 'scale(1)');
        }

        // tslint:disable-next-line:no-any
        private renderTable(data: any, index: number, categories: any, dataColumns: string[], options: any): void {
            $('.menuIsolate').hide();
            const thisObj: this = this;
            // sort method
            const gridSorter: (loop: number, iterator: number) => void = function (loop: number, iterator: number): void {

                // tslint:disable-next-line:no-any
                const tableContainer: any = d3.selectAll(`.bin${iterator}`);

                // tslint:disable-next-line:no-any
                const $this: any = this || `.tableHeader.no${loop}`;

                // tslint:disable-next-line:no-any
                const sortActive: any = tableContainer.selectAll('.sortActive');
                sortActive.select(`.sortIcon${iterator}`).style('visibility', 'hidden');
                sortActive.classed('sortActive', false);

                thisObj.sortedColumnIndex[index] = parseInt((d3.select($this).attr('sequence')), 10);

                tableContainer.select($this)
                    .classed('sortActive', thisObj.isSorted[index]);

                tableContainer.select('.sortActive').select(`.sortIcon${index}`)
                    .text(thisObj.isSortAsc[index] ? '▼' : '▲')
                    .style('visibility', 'visible');

                if (thisObj.isSortAsc[index]) {
                    thisObj.gridRow[iterator].sort(function (value1: DataViewTableRow, value2: DataViewTableRow): number {
                        if (value1[loop] === 'null') {
                            return -1;
                        } else if (value2[loop] === 'null') {
                            return 1;
                        }
                        // tslint:disable-next-line:no-any
                        const val1: any = isNaN(+value1[loop]) ? value1[loop] && value1[loop].toString().toLowerCase()
                            || '' : value1[loop] || '';
                        // tslint:disable-next-line:no-any
                        const val2: any = isNaN(+value2[loop]) ? value2[loop] && value2[loop].toString().toLowerCase()
                            || '' : value2[loop] || '';
                        const result: number = val1 < val2 ? 1 : -1;

                        return result;
                    });
                    thisObj.isSortAsc[index] = false;
                } else {
                    thisObj.gridRow[iterator].sort(function (value1: DataViewTableRow, value2: DataViewTableRow): number {
                        if (value1[loop] === 'null') {
                            return 1;
                        } else if (value2[loop] === 'null') {
                            return -1;
                        }
                        // tslint:disable-next-line:no-any
                        const val1: any = isNaN(+value1[loop]) ? value1[loop] && value1[loop].toString().toLowerCase()
                            || '' : value1[loop] || '';
                        // tslint:disable-next-line:no-any
                        const val2: any = isNaN(+value2[loop]) ? value2[loop] && value2[loop].toString().toLowerCase()
                            || '' : value2[loop] || '';
                        const result: number = val1 > val2 ? 1 : -1;

                        return result;
                    });
                    thisObj.isSortAsc[index] = true;
                }
            };

            // tslint:disable-next-line:no-any
            const rows: any[] = [];
            // tslint:disable-next-line:no-any
            let array: any[] = [];

            // tslint:disable-next-line:no-any
            let table: any;
            table = thisObj.mainCont
                .append('table')
                .classed(`tableCont bin${index}`, true);

            if (thisObj.numberCategory) {
                let caption: string = '';
                let limit: number;

                if (thisObj.selectionIndexes.length === 0) {
                    limit =
                        ((index * thisObj.numberOfBins) + thisObj.numberOfBins) > categories.length ?
                            categories.length : ((index * thisObj.numberOfBins) + thisObj.numberOfBins);
                    // No isolation;
                    if ((thisObj.numberOfBins > 1) && (categories[index * thisObj.numberOfBins] !== categories[limit - 1])) {
                        caption = `(${thisObj.binColumnformatter.format
                            (categories[index * thisObj.numberOfBins])} - ${thisObj.binColumnformatter.format(categories[limit - 1])})`;
                    } else {
                        caption = `(${thisObj.binColumnformatter.format(categories[index * thisObj.numberOfBins])})`;
                    }
                } else {
                    limit =
                        ((thisObj.selectedBins[index] * thisObj.numberOfBins) + thisObj.numberOfBins) > categories.length ?
                            categories.length : ((thisObj.selectedBins[index] * thisObj.numberOfBins) + thisObj.numberOfBins);

                    if ((thisObj.numberOfBins > 1) &&
                        (categories[thisObj.selectedBins[index] * thisObj.numberOfBins] !== categories[limit - 1])) {
                        caption = `(${(categories[thisObj.selectedBins[index] * thisObj.numberOfBins] === null ? 'Null' :
                            thisObj.binColumnformatter.format
                                (categories[thisObj.selectedBins[index] * thisObj.numberOfBins]))} - ${(categories[limit - 1] === null ?
                                    'Null' : thisObj.binColumnformatter.format
                                        (categories[limit - 1]))})`;
                    } else {
                        caption = `(${(categories[thisObj.selectedBins[index] * thisObj.numberOfBins] === null ? 'Null' :
                            thisObj.binColumnformatter.format(categories[thisObj.selectedBins[index] * thisObj.numberOfBins]))})`;
                    }
                }

                table.append('caption')
                    .style('font-weight', 'bold')
                    .style({
                        'font-size': `${this.settings.fontSettings.fontSize}px`,
                        'font-family': this.settings.fontSettings.fontFamily
                    })
                    .text(caption)
                    .attr('title', caption);
            } else {
                table.append('caption')
                    .style('font-weight', 'bold')
                    .style({
                        'font-size': `${this.settings.fontSettings.fontSize}px`,
                        'font-family': this.settings.fontSettings.fontFamily
                    })
                    .text(data[0][thisObj.binColumn] === null ? 'Null' : data[0][thisObj.binColumn])
                    .attr('title', data[0][thisObj.binColumn] === null ? 'Null' : data[0][thisObj.binColumn]);
            }

            table.append('thead').append('tr')
                .selectAll('th')
                .data(dataColumns).enter()
                .append('th')
                .style('border-bottom', `1px solid ${thisObj.colors[thisObj.iColumn][thisObj.jColumn]}`)
                .attr('class', function (datum: string, id: number): string {
                    return `tableHeader no${id}`;
                })
                .on('click', function (datum: string, id: number): void {
                    // i is column number ; index is bin number
                    thisObj.isSorted[index] = true;
                    gridSorter(id, index);
                })
                .append('p')
                .style('padding-right', '2px')
                .style('margin', '0px')
                .style('max-width', '150px')
                .style('white-space', 'nowrap')
                .style('overflow', 'hidden')
                .style('text-overflow', 'ellipsis')
                .style('color', 'black')
                // Align left if text, else align right for numbers
                .style('text-align', function (datum: string, id: number): string {
                    if (isNaN(+(data[0][thisObj.dataColIndex[id]]))) {
                        return 'left';
                    } else {
                        return 'right';
                    }
                })
                .style({
                    'font-size': `${this.settings.fontSettings.fontSize}px`,
                    'font-family': this.settings.fontSettings.fontFamily
                })
                .attr('class', function (datum: string, iterator: number): string {
                    return `column headercell header-${iterator}`;
                })
                .attr('sequence', function (datum: string, iterator: number): number {
                    return iterator;
                })
                .text(function (datum: string): string { return datum; })
                .attr('title', (function (datum: string): string { return datum; }));

            table.selectAll('.tableHeader')
                .append('span')
                .text(thisObj.isSortAsc[index] ? '▼' : '▲')
                .style('font-size', '8px')
                .style('color', 'black')
                .style('margin-left', '5px')
                .classed(`sortIcon${index}`, true)
                .style('float', 'left');

            d3.selectAll(`.sortIcon${index}`)
                .style('visibility', 'hidden');

            //Value Formatter Creation
            const formatter: IValueFormatter[] = [];
            for (let iterator: number = 0; iterator < thisObj.dataView.metadata.columns.length; iterator++) {
                const temp: IValueFormatter = ValueFormatter.create({
                    format: thisObj.dataView.metadata.columns[iterator].format
                });
                formatter.push(temp);
            }

            table.append('tbody')
                .selectAll('tr')
                .data(data).enter()
                .append('tr')
                .classed(`gridRow${index}`, true)
                .selectAll('td')
                // tslint:disable-next-line:no-any
                .data(function (datum: any): Object[] {
                    const arr: Object[] = [];
                    let cell: Object = {};
                    for (let iterator: number = 0; iterator < thisObj.dataColIndex.length; iterator++) {
                        cell = {};
                        cell[datum[thisObj.dataColIndex[iterator]]] =
                            (datum[thisObj.dataColIndex[iterator]] === null ? 'Null' :
                                formatter[thisObj.dataColIndex[iterator]].format(datum[thisObj.dataColIndex[iterator]]));
                        arr.push(cell);
                    }

                    return arr;
                })
                .enter()
                .append('td')
                .style({
                    'font-size': `${this.settings.fontSettings.fontSize}px`,
                    'font-family': this.settings.fontSettings.fontFamily
                })
                .style('max-width', '150px')
                .style('overflow', 'hidden')
                .style('text-overflow', 'ellipsis')
                .classed('gridcell', true)
                .style('color', thisObj.colors[thisObj.iColumntext][thisObj.jColumntext])
                .text(function (datum: Object): string {

                    // Creating array for sorting (non formatted values)
                    array.push(isNaN(Number(Object.keys(datum)[0])) ? //string or number
                        (isNaN(Date.parse(Object.keys(datum)[0])) ? Object.keys(datum)[0] :
                            Date.parse(Object.keys(datum)[0])) : //if string - date or not
                        (parseFloat(Object.keys(datum)[0])));

                    if (array.length === dataColumns.length) {
                        rows.push(array);
                        array = [];
                    }

                    // returning formatted values for display
                    return datum[Object.keys(datum)[0]];
                })
                // Text is aligned left and Numbers are aligned right
                .style('text-align', function (datum: Object): string {
                    if (isNaN(+(Object.keys(datum)[0]))) {
                        return 'left';
                    } else {
                        return 'right';
                    }
                })
                .attr({
                    class: 'rows',
                    title: function (datum: Object): string { return datum[Object.keys(datum)[0]]; }
                });

            thisObj.binData[index] = rows;
            thisObj.gridRow[index] = table.selectAll(`.gridRow${index}`).data(thisObj.binData[index]);
            thisObj.isSorted[index] = true;
            gridSorter(0, index);
            table.style('transform', 'scale(0)').transition().duration(500).style('transform', 'scale(1)');
        }
        // Render Legend
        private renderLegend1(): void {
            const thisObj: this = this;
            $('.legend1').remove();
            // tslint:disable-next-line:no-any
            let legend: any;
            if ((thisObj.maxLineIsOn || thisObj.minLineIsOn || thisObj.avgLineIsOn || thisObj.constantLineIsOn)
                && (thisObj.chartType === 'Bar' || thisObj.chartType === 'Column')) {
                legend = thisObj.mainCont.append('div')
                    .classed('legend1', true);
                if (thisObj.maxLineIsOn && thisObj.settings.legendLabel.show) {
                    legend.append('div')
                        .style('background-color', thisObj.maxLineFill)
                        .classed('maxLegend', true);
                    legend.append('div')
                        .text('Max')
                        .classed('maxNameLegend', true);
                }
                if (thisObj.minLineIsOn && thisObj.settings.legendLabel.show) {
                    legend.append('div')
                        .style('background-color', thisObj.minLineFill)
                        .classed('minLegend', true);
                    legend.append('div')
                        .text('Min')
                        .classed('minNameLegend', true);
                }
                if (thisObj.avgLineIsOn && thisObj.settings.legendLabel.show) {
                    legend.append('div')
                        .style('background-color', thisObj.avgLineFill)
                        .classed('avgLegend', true);
                    legend.append('div')
                        .text('Average')
                        .classed('avgNameLegend', true);
                }
                if (thisObj.constantLineIsOn && thisObj.settings.legendLabel.show) {
                    legend.append('div')
                        .style('background-color', thisObj.constantLineFill)
                        .classed('constantLegend', true);
                    legend.append('div')
                        .text('Constant')
                        .classed('constantNameLegend', true);
                }
            }
        }

        private calculateTableColumns(): string[] {
            const thisObj: this = this;
            const dataColumns: string[] = [];
            thisObj.dataColIndex = [];
            if (thisObj.selectedColumnsString === '') {
                return dataColumns;
            }
            const selectedColumnsStringArray: string[] = thisObj.selectedColumnsString.split('-', thisObj.viewModel.columns.length);
            for (let it: number = 0; it < selectedColumnsStringArray.length; it++) {
                dataColumns.push(thisObj.viewModel.columns[parseInt(selectedColumnsStringArray[it], 10)].displayName);
                thisObj.dataColIndex.push(parseInt(selectedColumnsStringArray[it], 10));
            }

            return dataColumns;
        }

        // tslint:disable-next-line:no-any
        private calculateDistinctBinValues(): any[] {
            const thisObj: this = this;
            // tslint:disable-next-line:no-any
            const arrayofbinvalues: any = [];
            for (let iterator: number = 0; iterator < thisObj.viewModel.dataPoints.length; iterator++) {
                if (arrayofbinvalues.indexOf(thisObj.viewModel.dataPoints[iterator].value[thisObj.binColumn]) === -1) {
                    arrayofbinvalues.push(thisObj.viewModel.dataPoints[iterator].value[thisObj.binColumn]);
                }
            }

            return arrayofbinvalues;
        }

        // Mapping data to table data for four cases : Isolation (Categorical, Numeric), No Isolation (categorical, Numeric)
        // tslint:disable-next-line:cyclomatic-complexity no-any
        private calculateTableData(binData: any, finalData: any, categories: string[]): any[] {
            const thisObj: this = this;
            // tslint:disable-next-line:no-any
            let tableData: any = [];
            // tslint:disable-next-line:no-any
            const tabledata1: any = [];
            // tslint:disable-next-line:no-any
            const arrayofbinvalues: any = thisObj.calculateDistinctBinValues();
            // No Isolation
            if (thisObj.selectionIndexes.length === 0) {
                for (let iterator: number = 0; iterator < binData.length; iterator++) {
                    // tslint:disable-next-line:no-any
                    const bins: any[] = [];
                    for (let innerIterator: number = 0; innerIterator < binData[iterator].values.length; innerIterator++) {
                        // tslint:disable-next-line:no-any
                        const values: any[] = [];
                        for (let dataIterator: number = 0; dataIterator < binData[iterator].values[innerIterator].value.length;
                            dataIterator++) {
                            values.push(binData[iterator].values[innerIterator].value[dataIterator]);
                        }
                        bins.push(values);
                    }
                    tableData.push(bins);
                }
            } else {
                // Isolation
                // tslint:disable-next-line:no-any
                const keysArray: any[] = [];
                for (let it: number = 0; it < finalData.length; it++) {
                    if (finalData[it] !== undefined) {
                        for (let iterator: number = 0; iterator < finalData[it].length; iterator++) {
                            keysArray.push(isNaN(Date.parse(finalData[it][iterator].key)) ?
                                finalData[it][iterator].key : Date.parse(finalData[it][iterator].key));
                        }
                    }
                }
                for (let outerIterator: number = 0; outerIterator < binData.length; outerIterator++) {
                    // tslint:disable-next-line:no-any
                    const bins: any[] = [];

                    let selectedBinIndex: number;

                    for (let innerIterator: number = 0; innerIterator < binData[outerIterator].values.length; innerIterator++) {
                        // tslint:disable-next-line:no-any
                        const values: any[] = [];
                        if (thisObj.numberCategory) {
                            selectedBinIndex = Math.floor(
                                categories.indexOf(binData[outerIterator].values[innerIterator].value[thisObj.binColumn])
                                / thisObj.numberOfBins);
                        }

                        if (keysArray.indexOf(binData[outerIterator].values[innerIterator].value[thisObj.keyColumnIndex] === null ? 'null' :
                            isNaN(Date.parse(binData[outerIterator].values[innerIterator].value[thisObj.keyColumnIndex])) ?
                                binData[outerIterator].values[innerIterator].value[thisObj.keyColumnIndex] :
                                Date.parse(binData[outerIterator].values[innerIterator].value[thisObj.keyColumnIndex])) > -1 &&
                            thisObj.selectedBins.indexOf(thisObj.numberCategory ? selectedBinIndex :
                                categories.indexOf(binData[outerIterator].values[innerIterator].value[thisObj.binColumn])) > -1) {
                            for (let dataIterator: number = 0; dataIterator < binData[outerIterator].values[innerIterator].value.length;
                                dataIterator++) {
                                values.push(binData[outerIterator].values[innerIterator].value[dataIterator]);
                            }
                            bins.push(values);
                        }
                    }
                    if (bins.length !== 0) {
                        tableData.push(bins);
                    }
                }
            }

            if (thisObj.numberCategory) {
                // Displaying in bin range
                // tslint:disable-next-line:no-any
                let temp: any[] = [];
                for (let iterator: number = 0; iterator < (thisObj.selectionIndexes.length === 0 ?
                    arrayofbinvalues.length : tableData.length); iterator++) {
                    if (iterator !== 0 && ((iterator % thisObj.numberOfBins) === 0)) {
                        tabledata1.push(temp);
                        temp = [];
                    }
                    for (let innerIterator: number = 0; innerIterator < tableData[iterator].length; innerIterator++) {
                        temp.push(tableData[iterator][innerIterator]);
                    }
                }
                tabledata1.push(temp);
                tableData = tabledata1;
            }

            return tableData;
        }

        private drawColumnSelector(dataColumns: string[]): void {
            const thisObj: this = this;
            thisObj.$xAxisLabel.hide();
            thisObj.$yAxisLabel.hide();
            $('.columnSelector').remove();
            $('.columnSelectorLabel').remove();
            //Column Selector
            // tslint:disable-next-line:no-any
            const columnSelectorLabel: any = thisObj.xAxisCont.append('p')
                .classed('columnSelectorLabel', true)
                .text(`Column selector ▲`)
                .style('display', 'table')
                .style('margin-left', function (): string {
                    const parentWidth: number = $(this).width();

                    return `${(thisObj.width / 2) - (parentWidth / 2) - 40}px`;
                });
            // tslint:disable-next-line:no-any
            const cont: any = thisObj.xAxisCont.append('div')
                .classed(`columnSelector`, true);

            thisObj.$columnSelector = $('.columnSelector');
            // tslint:disable-next-line:no-any
            const form: any = cont.append('form');
            // tslint:disable-next-line:no-any
            const columns: any = cont.selectAll('input')
                .data(thisObj.viewModel.columns).enter()
                .append('div')
                //tslint:disable-next-line:no-any
                .on('click', function (datum: any, indx: number): void {
                    const temp: JQuery = $(`input[data-id=${indx}]`);

                    $(temp).attr({ checked: !temp.attr('checked') });

                    if ($(temp).attr('checked') === 'checked') {

                        // If element is checked add to dataColumns
                        $(temp).prop('checked', true);
                        $(temp).attr('value', 'true');

                        dataColumns.push(temp.attr('name'));
                        thisObj.dataColIndex.push(indx);

                    } else {

                        // If element is unchecked remove from dataColumns
                        $(temp).attr('value', 'false');
                        $(temp).prop('checked', false);
                        // tslint:disable-next-line:no-any
                        const pop: any = function remove(array: any, element: any): any {
                            // tslint:disable-next-line:typedef
                            return array.filter(newElement => newElement !== element);
                        };

                        dataColumns = pop(dataColumns, temp.attr('name'));
                        thisObj.dataColIndex = pop(thisObj.dataColIndex, indx);
                    }
                    // Flag indicates that changes are made in column selector dropdown,
                    // hence now on hiding this dropdown, the chart must be redrawn
                    thisObj.flag = 1;
                })
                .classed('selector', true)
                // tslint:disable-next-line:no-any
                .text(function (datum: any, indx: number): string {
                    return datum.displayName;
                })
                .style('display', 'table')
                .append('input')
                .attr({
                    type: 'checkbox',
                    // tslint:disable-next-line:no-any
                    class: function (datum: any, indx: number): string {
                        return `dropdown.column.data${indx}`;
                    },
                    // tslint:disable-next-line:no-any
                    'data-id': function (datum: any, indx: number): any {
                        return indx;
                    },
                    value: 'false',
                    // tslint:disable-next-line:no-any
                    name: function (datum: any, indx: number): string {
                        return datum.displayName;
                    },
                    // tslint:disable-next-line:no-any
                    checked: function (datum: any, indx: number): any {
                        if (thisObj.dataColIndex.indexOf(indx) !== -1) {
                            return true;
                        }
                    }
                })
                // tslint:disable-next-line:no-any
                .text(function (datum: any): string { return `${datum.displayName}`; })
                .style('float', 'left');
            // Sorting the columns in X Axis dropdown
            thisObj.dropdown = cont.selectAll(`.selector`).data(thisObj.viewModel.columns);
            thisObj.columnSorter();

            d3.select('.columnSelector').style('margin-left', function (): string {
                const parentWidth: number = $(this).width();

                return `${(thisObj.width / 2) - (parentWidth / 2) - 40}px`;
            });

            columnSelectorLabel.on('click', function (): void {
                thisObj.hideMenus(thisObj.$columnSelector);
            });
        }
        // tslint:disable-next-line:no-any
        private createBinData(categories: any): any {
            const thisObj: this = this;
            // tslint:disable-next-line:no-any
            const binData: any = d3.nest()
                // tslint:disable-next-line:no-any
                .key(function (datum: IDataPoints): any {
                    return datum[`value`][thisObj.binColumn];
                })
                // tslint:disable-next-line:no-any
                .rollup(function (datum: any): any {
                    categories.push(datum[0][`value`][thisObj.binColumn]);

                    return datum;
                }).entries(thisObj.viewModel.dataPoints);

            return binData;
        }

        // tslint:disable-next-line:no-any
        private createcolorData(colorby: any): any {

            const thisObj: this = this;
            // tslint:disable-next-line:no-any
            const colorData: any = d3.nest()
                // tslint:disable-next-line:no-any
                .key(function (datum: IDataPoints): any {
                    return datum[`value`][thisObj.colorByColumn];
                })
                // tslint:disable-next-line:no-any
                .rollup(function (datum: any): any {
                    colorby.push(datum[0][`value`][thisObj.colorByColumn]);

                    return datum;
                }).entries(thisObj.viewModel.dataPoints);

            return colorData;
        }

        private getHeight(currentHeight: number, countOfFirstBox: number): number {

            if (currentHeight <= (this.mainContHeight / countOfFirstBox)) {

                return $('.chart').height() - 20;
            }

            return (this.mainContHeight / countOfFirstBox -
                (60 / countOfFirstBox)) - 20;
        }

        private returnSkip(chartWidth: number): number {
            if ((chartWidth + 12) * 3 < this.mainContWidth) {
                return 3;

            } else if ((chartWidth + 12) * 2 < this.mainContWidth) {
                return 2;
            } else {
                return 1;
            }
        }

        private hideWhenIsolate(): void {
            const thisObj: this = this;
            if (thisObj.selectionIndexes.length > 0) {
                thisObj.$xAxisLabel.hide();
                thisObj.$yAxisLabel.hide();
                $('.menuX').hide();
                $('.menuY').hide();
                $('.menuBinningby').hide();
                thisObj.$xAxis.hide();
                thisObj.$yAxis.hide();
                thisObj.$binningCont.hide();

                d3.select('.colorCont.shape').style('left', $('.menuColor').position().left + thisObj.px);
                d3.select('.colorCont.text').style('left', $('.menuTextColor').position().left + thisObj.px);
            }
        }

        // tslint:disable-next-line:typedef
        private getRandomColor() {
            const letters: string = '0123456789ABCDEF';
            let color: string = '#';
            // tslint:disable-next-line:typedef
            for (let getcolor = 0; getcolor < 6; getcolor++) {
                color += letters[Math.floor(Math.random() * 16)];
            }

            return color;
        }

        // tslint:disable-next-line:cyclomatic-complexity
        private renderChart(): void {
            const thisObj: this = this;

            // Formatter values
            thisObj.binColumnformatter = ValueFormatter.create({
                format: (thisObj.binColumn !== -1 ? thisObj.dataView.metadata.columns[thisObj.binColumn].format : '')
            });

            thisObj.targetColumnformatter = ValueFormatter.create({
                format: thisObj.dataView.metadata.columns[thisObj.targetColumn].format
            });
            if (!thisObj.settings.presentation.show) {
                thisObj.renderMenu();
            } else {
                thisObj.renderMenu();
                d3.select('.topMenu').remove();
                d3.select('.topCont').remove();
            }
            d3.selectAll('.tableCont,.chart,.brickChart, .columnChart, .brickChart1, .borderChart').remove();
            let dataColumns: string[] = [];
            dataColumns = thisObj.calculateTableColumns();

            if (0 === thisObj.actions.length) {
                $('.undoCont').hide();
            } else {
                $('.undoCont').show();
            }
            let legendData: LegendData;
            //  let dataView: DataView;
            legendData = {
                fontSize: 9,
                dataPoints: []
            };
            // tslint:disable-next-line:no-any
            const chartData: any = [];
            thisObj.selectedBins = [];
            let category: string[] = [];
            if (thisObj.chartType !== thisObj.prevChartType || thisObj.updateCalled) {
                thisObj.globalSelections = [];
                let iCounter: number = 0;
                //;
                for (iCounter = 0; iCounter < thisObj.prevGlobalSelections.length; iCounter++) {
                    thisObj.globalSelections[iCounter] = thisObj.prevGlobalSelections[iCounter];
                }
            }
            thisObj.updateCalled = false;
            if (thisObj.undoPressed && thisObj.globalSelectionsOld.length > 0) {
                let iCounter: number;
                for (iCounter = 0; iCounter < thisObj.globalSelectionsOld[thisObj.globalSelectionsOld.length - 1].length; iCounter++) {
                    thisObj.globalSelections[iCounter] = thisObj.globalSelectionsOld[thisObj.globalSelectionsOld.length - 1][iCounter];
                }
            }
            thisObj.undoPressed = false;
            // tslint:disable-next-line:no-any
            thisObj.selectionIndexes.forEach(function (index: number): any {
                if (chartData[thisObj.globalSelections[index].binIndex]) {
                    chartData[thisObj.globalSelections[index].binIndex].push(thisObj.globalSelections[index].data);
                } else {
                    chartData[thisObj.globalSelections[index].binIndex] = [thisObj.globalSelections[index].data];
                }
                if (thisObj.selectedBins.indexOf(thisObj.globalSelections[index].binIndex) === -1) {
                    thisObj.selectedBins.push(thisObj.globalSelections[index].binIndex);
                }
                category.push(thisObj.globalSelections[index].category);

                // tslint:disable-next-line:no-any
                return thisObj.globalSelections[index];
            });
            // tslint:disable-next-line:no-any
            category = category.filter(function (item: any, pos: number): any {
                return category.indexOf(item) === pos;
            });
            thisObj.prevChartType = thisObj.chartType;
            let iCounter1: number;
            iCounter1 = 0;
            for (iCounter1 = 0; iCounter1 < thisObj.globalSelections.length; iCounter1++) {
                thisObj.prevGlobalSelections[iCounter1] = thisObj.globalSelections[iCounter1];
            }
            const categories: string[] = [];
            const colorby: string[] = [];
            // tslint:disable-next-line:no-any
            const binData: any = thisObj.createBinData(categories);
            // tslint:disable-next-line:no-any
            const colorData: any = thisObj.createcolorData(colorby);

            // Limit on number of rows
            if (thisObj.viewModel.dataPoints.length > 20000) {
                thisObj.mainCont.append('div')
                    .classed('tableCont', true)
                    .text('Current visual supports up to 20K rows. Please filter the dataset and reduce the number of rows.')
                    .style('margin-top', '40px');
                d3.select('.label0').text('None');

                return;
            }

            // Limit on unique values of binning category
            if (categories.length > 1200) {
                thisObj.mainCont.append('div')
                    .classed('tableCont', true)
                    .text('Please use a different binning category as this one has many unique values')
                    .style('margin-top', '40px');
                d3.select('.label0').text('None');

                return;
            }

            const numberString: string = 'number';
            // tslint:disable-next-line:no-any
            let finalData: any = [];
            thisObj.numberCategory = typeof categories[0] === numberString ? true : false;
            thisObj.mainContWidth = thisObj.$mainCont.width();
            thisObj.mainContHeight = (thisObj.$mainCont.height());
            thisObj.renderLegend1();

            thisObj.globalSelections = [];
            // tslint:disable-next-line:no-any
            let k: any = [];
            // tslint:disable-next-line:no-any
            const colorByValues: any = [];
            // Facets
            if (thisObj.selectionIndexes.length === 0) {
                // tslint:disable-next-line:no-any
                binData.forEach(function (data: any): any {
                    finalData.push(d3.nest()
                        // tslint:disable-next-line:no-any
                        .key(function (datum: any): any {
                            thisObj.keyColumnIndex = thisObj.groupedColumn;

                            return datum[`value`][thisObj.groupedColumn];
                        })
                        // tslint:disable-next-line:no-any
                        .rollup(function (data1: any): any {
                            // Again nest to get the values according to the color by values
                            if (thisObj.colorByColumn !== -1) {
                                k = [];
                                k.push(d3.nest()
                                    // tslint:disable-next-line:no-any
                                    .key(function (dat: any): any {
                                        colorByValues.push(dat[`value`][thisObj.colorByColumn]);

                                        return dat[`value`][thisObj.colorByColumn];
                                    })
                                    // tslint:disable-next-line:no-any
                                    .rollup(function (dat: any): any {
                                        let value: number;
                                        // tslint:disable-next-line:no-any
                                        value = d3.sum(dat, function (dataIterator: any): number {

                                            if (dataIterator[`value`][thisObj.targetColumn] === null ||
                                                isNaN(Number(dataIterator[`value`][thisObj.targetColumn]))) {

                                                return 0;
                                            } else {

                                                return dataIterator[`value`][thisObj.targetColumn];
                                            }
                                        });
                                        if (value > 0) {
                                            thisObj.isCategory = false;
                                        }
                                        // tslint:disable-next-line:no-any
                                        const selectionIds: any = [];
                                        // tslint:disable-next-line:no-any
                                        dat.forEach(function (innerdata: any): void {
                                            selectionIds.push(innerdata[`selectionId`]);
                                        });

                                        return {
                                            value: value,
                                            selectionId: selectionIds
                                        };
                                    }).entries(data1));

                                return k;
                            } else {
                                let value: number;
                                // tslint:disable-next-line:no-any
                                value = d3.sum(data1, function (dataIterator: any): number {

                                    if (dataIterator[`value`][thisObj.targetColumn] === null ||
                                        isNaN(Number(dataIterator[`value`][thisObj.targetColumn].toString()))) {

                                        return 0;
                                    } else {

                                        return dataIterator[`value`][thisObj.targetColumn];
                                    }
                                });
                                if (value > 0) {
                                    thisObj.isCategory = false;
                                }
                                // tslint:disable-next-line:no-any
                                const selectionIds: any = [];
                                // tslint:disable-next-line:no-any
                                data1.forEach(function (innerdata: any): void {
                                    selectionIds.push(innerdata[`selectionId`]);
                                });

                                return {
                                    value: value,
                                    selectionId: selectionIds
                                };
                            }
                        }).entries(data.values));
                });
            } else {
                finalData = chartData;
            }
            if (thisObj.counter === 0) {
                thisObj.previousColorByColumn = thisObj.colorByColumn;
                thisObj.barcolor = [];
                // tslint:disable-next-line:typedef
                colorby.forEach(element => {
                    thisObj.barcolor.push({ key: element == null ? 'null' : element, value: this.getRandomColor() });
                });
            }
            if (thisObj.previousColorByColumn !== thisObj.colorByColumn || thisObj.counter === 0) {
                thisObj.counter++;
                thisObj.previousColorByColumn = thisObj.colorByColumn;
                thisObj.barcolor = [];
                // tslint:disable-next-line:typedef
                colorby.forEach(element => {
                    thisObj.barcolor.push({ key: element == null ? 'null' : element, value: this.getRandomColor() });
                });
            }

            for (let legendIterator: number = 0; legendIterator < this.barcolor.length; legendIterator++) {
                legendData.dataPoints.push({
                    label: this.barcolor[legendIterator].key.toString(),
                    color: this.barcolor[legendIterator].value,
                    icon: powerbi.extensibility.utils.chart.legend.LegendIcon.Box,
                    selected: false,
                    identity: this.host.createSelectionIdBuilder().withCategory(
                        this.currentDataview.categorical.categories[0], legendIterator).createSelectionId()
                });
            }
            d3.select(`.legend`)
                .style('height', '22px');
            const position: LegendPosition = LegendPosition.Top;
            this.legend.changeOrientation(position);
            if (thisObj.settings.legendLabel.show && this.colorByColumn !== -1) {

                thisObj.mainCont.style('padding-top', '20px');
                $('.legend').show();
                this.legend.drawLegend(legendData, this.currentViewport);

                d3.select('.legend').style('overflow', 'auto');
            } else {
                thisObj.mainCont.style('padding-top', '0px');
                $('.legend').hide();

            }
            if (thisObj.chartType.toLowerCase() === 'table') {
                $('.legend').hide();
            }
            // get the unique values of color by column
            // tslint:disable-next-line:no-any
            const uniqueColorByValues: any = colorByValues.filter(function (item: any, pos: any): any {
                return colorByValues.indexOf(item) === pos;
            });
            thisObj.colorByValuesLength = uniqueColorByValues.length;
            const tabwidth: number = 95; //presentation tab width
            let totalCategories: number;
            const margin: number = 30;
            totalCategories = thisObj.selectionIndexes.length === 0 ? categories.length : category.length;
            let chartWidth: number = totalCategories <= 2 ? Math.max(Math.floor(thisObj.mainContWidth /
                totalCategories) - margin,                           350)
                : Math.max(Math.floor((thisObj.mainContWidth / 3)) - margin, 350);
            let skip: number;
            skip = thisObj.returnSkip(chartWidth);
            let countOfFirstBox: number = 0;
            let indexCounter: number = 0;
            const isLegend: number = (thisObj.maxLineIsOn || thisObj.minLineIsOn || thisObj.avgLineIsOn || thisObj.constantLineIsOn)
                ? $('.legend1').height() : 0;
            const chartHeight: number = totalCategories <= skip ? thisObj.mainContHeight - $('.topCont').height() - isLegend :
                Math.floor(thisObj.mainContHeight / 2) - 30;
            //Checks if bin and  undo button is displayed and accordingly calculate  no. of menu items and sets width for each li in  %
            $('.menuRangeforBinning').hide();
            if (0 === thisObj.actions.length) {
                this.countOfMenuItems = thisObj.countOfMenuItems - 2;
                //min from 95 as 5% goes for margin
                d3.selectAll('li')
                    .style('width', `${Math.floor(tabwidth / thisObj.countOfMenuItems)}` + '%');
            } else {
                this.countOfMenuItems = thisObj.countOfMenuItems - 1;
                //min from 95 as 5% goes for margin
                d3.selectAll('li')
                    .style('width', `${Math.floor(tabwidth / thisObj.countOfMenuItems)}` + '%');
            }
            if (thisObj.numberCategory) {
                //Checks if bin and  undo button is displayed and accordingly calculate  no. of menu items and sets width for each li in  %
                $('.menuRangeforBinning').show();
                if (0 === thisObj.actions.length) {
                    this.countOfMenuItems = thisObj.countOfMenuItems - 1;
                    //min from 95 as 5% goes for margin
                    d3.selectAll('li')
                        .style('width', `${Math.floor(tabwidth / thisObj.countOfMenuItems)}` + '%');
                } else {
                    this.countOfMenuItems = thisObj.countOfMenuItems;
                    //min from 95 as 5% goes for margin
                    d3.selectAll('li')
                        .style('width', `${Math.floor(tabwidth / thisObj.countOfMenuItems)}` + '%');
                }
                // tslint:disable-next-line:prefer-const
                let countValues: {};
                const binSize: number = finalData.length % thisObj.numberOfBins === 0 ?
                    finalData.length / thisObj.numberOfBins : Math.floor(finalData.length / thisObj.numberOfBins);
                let counter: number;
                let count: number;
                // tslint:disable-next-line:no-any
                let binRangeData: any = [];
                // tslint:disable-next-line:no-any
                let catData: any = [];
                let iCounter: number;
                let jCounter: number;
                if (thisObj.selectionIndexes.length > 0) {
                    $('.menuRangeforBinning').hide();
                    catData = category;
                    counter = 0;
                    let lCounterNew: number = 0;
                    let prevDataLengthNew: number = 0;
                    thisObj.previousDataLength = 0;
                    totalCategories = catData.length;
                    chartWidth = totalCategories <= 2 ? Math.max(Math.floor(thisObj.mainContWidth /
                        totalCategories) - margin,               350)
                        : Math.max(Math.floor((thisObj.mainContWidth / 3)) - margin, 350);
                    skip = thisObj.returnSkip(chartWidth);
                    // tslint:disable-next-line:no-any
                    finalData.forEach(function (data: any, index: number): void {
                        thisObj.$xAxisLabel.hide();
                        thisObj.$yAxisLabel.hide();
                        $('.menuX').hide();
                        $('.menuY').hide();
                        $('.menuBinningby').hide();
                        thisObj.$xAxis.hide();
                        thisObj.$yAxis.hide();
                        thisObj.$binningCont.hide();

                        d3.select('.colorCont.shape').style('left', $('.menuColor').position().left + thisObj.px);
                        d3.select('.colorCont.text').style('left', $('.menuTextColor').position().left + thisObj.px);
                        if (thisObj.chartType.toLowerCase() === 'bar' || thisObj.chartType.toLowerCase() === 'brick') {
                            if (finalData[index] === undefined) {
                                return;
                            }
                            if (counter % skip === 0) {
                                countOfFirstBox++;
                                totalData = 0;
                                let rowIndex: number;
                                for (rowIndex = counter; rowIndex < counter + skip; rowIndex++) {
                                    if (rowIndex < finalData.length) {
                                        if (finalData[rowIndex] !== undefined) {
                                            totalData = totalData + finalData[rowIndex].length;
                                        }
                                    }
                                }
                                totalData = totalData / skip;
                                // 15 is bar height; 16 is label Height; 10 is label margin-bottom
                                if (thisObj.chartType.toLowerCase() === 'bar') {
                                    totalData = Math.min((totalData * 15) + 16 + 10, thisObj.mainContHeight);
                                }
                                if (thisObj.chartType.toLowerCase() === 'brick') {
                                    totalData = Math.min(((totalBinData / Math.floor(chartWidth / 60)) * thisObj.textSize * 2) + 16 + 10,
                                                         thisObj.mainContHeight);
                                }
                            }
                        }
                        if (categories.length <= 3 && thisObj.chartType.toLowerCase() === 'bar') {
                            totalData = thisObj.mainContHeight - 50;
                        }

                        thisObj.previousDataLength = lCounterNew === 0 ? 0 : thisObj.previousDataLength + prevDataLengthNew;
                        lCounterNew++;
                        prevDataLengthNew = data.length;

                        const marginForCenter: number = thisObj.returnMarginForCenter(counter, skip, chartWidth);
                        if (thisObj.chartType.toLowerCase() === 'bar') {
                            thisObj.renderBinBarChart(data, counter, catData, chartWidth, chartHeight, totalData,
                                                      marginForCenter);
                        } else if (thisObj.chartType.toLowerCase() === 'brick') {
                            thisObj.renderBinBrickChart(data, counter, catData, chartWidth, totalData,
                                                        marginForCenter);
                        } else if (thisObj.chartType.toLowerCase() === 'column') {
                            thisObj.renderBinColumnChart(data, counter, catData, chartWidth, chartHeight,
                                                         marginForCenter);
                        }
                        counter++;
                    });
                } else {
                    // tslint:disable-next-line:max-line-length
                    for (counter = 0, count = thisObj.numberOfBins; finalData.length % thisObj.numberOfBins === 0 ? counter < binSize : counter <= binSize; count = count + thisObj.numberOfBins, counter++) {

                        // tslint:disable-next-line:no-any
                        const array: any[] = [];
                        // tslint:disable-next-line:no-any
                        const vals: any[] = [];
                        // tslint:disable-next-line:no-any
                        const keys: any = [];
                        // tslint:disable-next-line:no-any
                        const valKeys: any = [];
                        // tslint:disable-next-line:no-any
                        let uniqueKeys: any;
                        // tslint:disable-next-line:no-any
                        let uniqueValKeys: any;
                        uniqueValKeys = [];
                        uniqueKeys = [];
                        const sum: number[] = [];
                        const selectionIDs: ISelectionId[] = [];
                        // tslint:disable-next-line:no-any
                        let abc: any[];
                        abc = [];
                        // tslint:disable-next-line:no-any
                        finalData.forEach(function (data: any, index: number): any {
                            if (index >= count - thisObj.numberOfBins && index < count) {
                                // tslint:disable-next-line:no-any
                                data.forEach(function (datum: any): any {
                                    if (datum.key !== '') {
                                        array.push(datum);
                                        if (thisObj.colorByColumn !== -1) {
                                            vals.push(datum[`values`]);
                                        }
                                        keys.push(datum.key);
                                        selectionIDs.push(datum.values[`selectionId`]);
                                    }
                                });
                            }

                            abc.push(data);

                        });
                        // tslint:disable-next-line:no-any
                        uniqueKeys = keys.filter(function (item: any, pos: any): any {
                            return keys.indexOf(item) === pos;
                        });
                        let sumCount: number;
                        if (thisObj.colorByColumn === -1) {
                            for (iCounter = 0; iCounter < uniqueKeys.length; iCounter++) {
                                sumCount = 0;
                                for (jCounter = 0; jCounter < array.length; jCounter++) {
                                    if (uniqueKeys[iCounter] === array[jCounter].key) {
                                        sumCount = sumCount + array[jCounter].values[`value`];
                                    }
                                }
                                sum.push(sumCount);
                            }
                            // tslint:disable-next-line:no-any
                            countValues = uniqueKeys.map(function (element: any, index: any): any {
                                return { key: element, values: { value: sum[index], selectionId: selectionIDs } };
                            });
                        } else {
                            //do nothing
                        }

                        const secondName: string = (counter === binSize ?
                            categories[categories.length - 1] : categories[count - 1]);
                        const cat: string = thisObj.numberOfBins === 1 ? categories[count - thisObj.numberOfBins] :
                            `(${(categories[count - thisObj.numberOfBins] === null ? 'Null' :
                                thisObj.binColumnformatter.format(categories[count - thisObj.numberOfBins]))} - ${secondName === null ?
                                    'Null' : thisObj.binColumnformatter.format(secondName)})`;
                        binRangeData.push(abc[counter]);
                        catData.push(cat);

                    }
                }
                // CHECK FOR DATA

                // tslint:disable-next-line:no-any
                binRangeData = binRangeData.map(function (element: any, index: number): any {
                    return {

                        data: element,
                        cat: catData[index]
                    };
                });
                let lCounter: number;
                lCounter = 0;
                thisObj.previousDataLength = 0;
                let totalBinData: number;
                totalBinData = 0;
                totalCategories = catData.length;
                chartWidth = totalCategories <= 2 ? Math.max(Math.floor(thisObj.mainContWidth /
                    totalCategories) - margin,               350)
                    : Math.max(Math.floor((thisObj.mainContWidth / 3)) - margin, 350);
                skip = thisObj.returnSkip(chartWidth);
                // tslint:disable-next-line:no-any
                binRangeData.forEach(function (data: any, index: number): any {
                    if (index % skip === 0) {
                        countOfFirstBox++;
                        totalBinData = 0;
                        let rowIndex: number;
                        for (rowIndex = index; rowIndex < index + skip; rowIndex++) {
                            if (rowIndex < binRangeData.length) {
                                totalBinData = totalBinData + binRangeData[rowIndex].data.length;
                            }
                        }
                        totalBinData = totalBinData / skip;

                        // 15 is bar height; 16 is label Height; 10 is label margin-bottom
                        if (thisObj.chartType.toLowerCase() === 'bar') {
                            totalBinData = Math.min((totalBinData * 15) + 16 + 10, thisObj.mainContHeight);
                        }
                        if (thisObj.chartType.toLowerCase() === 'brick') {
                            totalBinData = Math.min(((totalBinData / Math.floor(chartWidth / 60)) * 35) + 16 + 10, thisObj.mainContHeight);
                        }
                    }
                    thisObj.previousDataLength = lCounter === 0 ? 0 : thisObj.previousDataLength + binRangeData[lCounter - 1].data.length;
                    lCounter++;

                    const marginForCenter: number = thisObj.returnMarginForCenter(index, skip, chartWidth);
                    if ('bar' === thisObj.chartType.toLowerCase()) {
                        thisObj.renderBinBarChart(data.data, index, catData, chartWidth, chartHeight, totalBinData,
                                                  marginForCenter);
                    }
                    if ('brick' === thisObj.chartType.toLowerCase()) {

                        thisObj.renderBinBrickChart(data.data, index, catData, chartWidth, totalBinData,
                                                    marginForCenter);
                    }
                    if ('column' === thisObj.chartType.toLowerCase()) {
                        thisObj.renderBinColumnChart(data.data, index, catData, chartWidth, chartHeight,
                                                     marginForCenter);
                    }
                });
            }
            //Dynamic placement of dropdowns due to change in assignment of width of topmenu bar
            const dropDownRealignment: () => void = function (): void {
                if (undefined !== $('.topMenu')[0]) {
                    let menuLeftPosition: JQuery = $('.menuBinningby');
                    $('.binningCont').css({ left: menuLeftPosition.position().left + thisObj.px });
                    // tslint:disable-next-line:no-any
                    menuLeftPosition = $('.menuX');
                    $('.targetCont').css({
                        left: menuLeftPosition.position().left + thisObj.px,
                        top: menuLeftPosition.position().top + menuLeftPosition.height() + thisObj.px
                    });

                    menuLeftPosition = $('.menuY');
                    $('.groupingCont').css({
                        left: menuLeftPosition.position().left + thisObj.px,
                        top: menuLeftPosition.position().top + menuLeftPosition.height() + thisObj.px
                    });

                    menuLeftPosition = $('.menuColor');
                    $('.shape').css({
                        left: menuLeftPosition.position().left + thisObj.px,
                        top: menuLeftPosition.position().top + menuLeftPosition.height() + thisObj.px
                    });
                    menuLeftPosition = $('.menuTextColor');
                    $('.text').css({
                        left: menuLeftPosition.position().left + thisObj.px,
                        top: menuLeftPosition.position().top + menuLeftPosition.height() + thisObj.px
                    });

                    menuLeftPosition = $('.menuRangeforBinning');
                    $('.dynamicBinCont').css({
                        left: menuLeftPosition.position().left + thisObj.px,
                        top: menuLeftPosition.position().top + menuLeftPosition.height() + thisObj.px
                    });

                    menuLeftPosition = $('.menuColorBy');
                    $('.colorByCont').css({
                        left: menuLeftPosition.position().left + thisObj.px,
                        top: menuLeftPosition.position().top + menuLeftPosition.height() + thisObj.px
                    });
                }
            };
            if ('table' !== thisObj.chartType.toLowerCase()) {
                // Below condition checks if the "Color By" option has "None" value
                if (thisObj.colorByColumn !== -1) {
                    //it hides the Color <li> when "Color By" is Selected other than "None"
                    $('.menuColor').hide();
                    dropDownRealignment();
                } else {
                    dropDownRealignment();
                }
            } else {
                /* we are not calling dropDonRealignment function here, as 'Color By' is the last menu that has dropdown.
                Hence we need not have to realign the dropdown menus of other options*/
                $('.menuColorBy').hide();
            }
            if (!thisObj.canUndo) {
                thisObj.canUndo = true;
                thisObj.cacheSelectionState();
            }
            let length: number;
            const categoriesLength: number = categories.length;
            const colorbyLength: number = colorby.length;
            let totalData: number = 0;
            let kCounter: number = 0;
            let prevDataLength: number = 0;
            thisObj.previousDataLength = 0;
            thisObj.colorByDataSelection = 0;
            if (!thisObj.numberCategory && thisObj.chartType.toLowerCase() !== 'none') {
                thisObj.hideWhenIsolate();
                // tslint:disable-next-line:no-any
                finalData.forEach(function (data: any, index: number): {} {
                    if (thisObj.chartType.toLowerCase() === 'bar' || thisObj.chartType.toLowerCase() === 'brick') {
                        if (finalData[index] === undefined) {
                            return;
                        }
                        if (indexCounter % skip === 0) {
                            countOfFirstBox++;
                            totalData = 0;
                            let rowIndex: number;
                            for (rowIndex = indexCounter; rowIndex < indexCounter + skip; rowIndex++) {
                                if (rowIndex < finalData.length) {
                                    if (finalData[rowIndex] !== undefined) {
                                        totalData = totalData + finalData[rowIndex].length;
                                    }
                                }
                            }

                            totalData = totalData / skip;
                            // 15 is bar height; 16 is label Height; 10 is label margin-bottom
                            if (thisObj.chartType.toLowerCase() === 'bar') {
                                totalData = Math.min((totalData * 15) + 16 + 10, thisObj.mainContHeight);
                            }
                            if (thisObj.chartType.toLowerCase() === 'brick') {
                                totalData = Math.min(((totalData / Math.floor(chartWidth / 60)) *
                                    (thisObj.textSize * 2) + 20) + 16 + 10,
                                                     thisObj.mainContHeight);
                            }
                        }
                    }
                    if (categories.length <= 3 && thisObj.chartType.toLowerCase() === 'bar') {
                        totalData = thisObj.mainContHeight - 50;
                    }
                    if (thisObj.colorByColumn === -1) {
                        if (thisObj.selectionIndexes.length === 0) {
                            thisObj.previousDataLength = kCounter === 0 ? 0 : thisObj.previousDataLength + finalData[kCounter - 1].length;
                        } else {
                            thisObj.previousDataLength = kCounter === 0 ? 0 : thisObj.previousDataLength + prevDataLength;
                            prevDataLength = data.length;
                        }
                    }
                    kCounter++;
                    const marginForCenter: number = thisObj.returnMarginForCenter(indexCounter, skip, chartWidth);
                    /***************BRICK CHART*******************/
                    if ('brick' === thisObj.chartType.toLowerCase()) {
                        length = data.length;
                        thisObj.previousLength = thisObj.previousLength < length
                            ? length : thisObj.previousLength;
                        thisObj.renderBrickChart(data, index, categories, length, chartWidth, totalData, marginForCenter);
                    }
                    /***************HTML BAR CHART*******************/
                    if ('bar' === thisObj.chartType.toLowerCase()) {
                        thisObj.renderBarChart(data, index, categories, chartWidth, chartHeight, totalData, marginForCenter);
                    }

                    /*********** COLUMN CHART *************/
                    if ('column' === thisObj.chartType.toLowerCase()) {
                        thisObj.renderColumnChart(data, index, categories, chartWidth, chartHeight, marginForCenter);
                    }
                    indexCounter++;
                });
            }
            if (thisObj.chartType.toLowerCase() === 'bar') {
                if (thisObj.mainContHeight > (countOfFirstBox * chartHeight) - skip * 10) {
                    d3.selectAll('.chart').style('height', thisObj.mainContHeight / countOfFirstBox - (60 / countOfFirstBox) + thisObj.px);
                    // tslint:disable-next-line:no-any
                    d3.selectAll('.subDiv').each(function (datum: any, index: number): void {
                        const currentHeight: number = $(this).height();
                        d3.select(this).classed(`subDiv${index}`);
                        d3.select(`.subDiv${index}`).style('height', thisObj.getHeight(currentHeight, countOfFirstBox) + thisObj.px);
                    }
                    );
                }
            }
            if (thisObj.chartType.toLowerCase() === 'brick') {
                if (thisObj.mainContHeight > (countOfFirstBox * chartHeight) - skip * 10) {
                    d3.selectAll('.brickChart').style('height', thisObj.mainContHeight / countOfFirstBox -
                        (60 / countOfFirstBox) + thisObj.px);
                }
            }

            /**************TABLE********************/
            if ('table' === thisObj.chartType.toLowerCase()) {
                thisObj.drawColumnSelector(dataColumns);
                // tslint:disable-next-line:no-any
                const tableData: any = thisObj.calculateTableData(binData, finalData, categories);
                // tslint:disable-next-line:no-any
                if (dataColumns.length !== 0) {
                    // tslint:disable-next-line:no-any
                    tableData.forEach(function (data: any, index: number): any {
                        thisObj.renderTable(data, index, categories, dataColumns, thisObj.options);
                    });
                }
            }
            if (thisObj.binColumn === -1 && thisObj.chartType.toLowerCase() !== 'bar') {
                $('.label').hide();
            }
            d3.selectAll('.legend').style('margin-left', '0px');
           // .style('margin-top', '40px');
        }

        //Use persisted properties
        public setOrient(options: VisualUpdateOptions): void {
            let objects: DataViewObjects = null;
            objects = this.dataView.metadata.objects;
            const chartTypeValue: string = this.settings.orient.prop;
            const xAxisValue: number = this.settings.orient.xaxis;
            const yAxisValue: number = this.settings.orient.yaxis;
            const binColumnValue: number = this.settings.orient.bin;
            const icolorValue: number = this.settings.orient.icolor;
            const jcolorValue: number = this.settings.orient.jcolor;
            const icolorValuetext: number = this.settings.orient.icolortext;
            const jcolorValuetext: number = this.settings.orient.jcolortext;
            const numberOfBins: number = this.settings.orient.noOfBins;
            const selectedColumns: string = this.settings.orient.columns;
            const colorByColumn: number = this.settings.orient.colorBy;

            this.chartType = String(chartTypeValue);
            this.groupedColumn = yAxisValue;
            this.targetColumn = xAxisValue;
            this.binColumn = binColumnValue;
            this.iColumn = icolorValue;
            this.jColumn = jcolorValue;
            this.iColumntext = icolorValuetext;
            this.jColumntext = jcolorValuetext;
            this.numberOfBins = numberOfBins;
            this.colorByColumn = colorByColumn;
            this.selectedColumnsString = selectedColumns;
        }

        //Persist Properties
        public persistOrient(): void {
            let thisObj: this;
            thisObj = this;
            let properties: { [propertyName: string]: DataViewPropertyValue };
            properties = {};
            properties[`prop`] = thisObj.chartType;
            properties[`yaxis`] = thisObj.groupedColumn;
            properties[`xaxis`] = thisObj.targetColumn;
            properties[`bin`] = thisObj.binColumn;
            properties[`icolor`] = thisObj.iColumn;
            properties[`jcolor`] = thisObj.jColumn;
            properties[`icolortext`] = thisObj.iColumntext;
            properties[`jcolortext`] = thisObj.jColumntext;
            properties[`noOfBins`] = thisObj.numberOfBins;
            properties[`columns`] = thisObj.selectedColumnsString;
            properties[`colorBy`] = thisObj.colorByColumn;
            let orient: VisualObjectInstancesToPersist;
            orient = {
                replace: [
                    <VisualObjectInstance>{
                        objectName: 'orient',
                        selector: null,
                        properties: properties
                    }]
            };
            this.host.persistProperties(orient);
        }

        public setColumns(column: string, value: number): void {
            const thisObj: this = this;
            if (column === 'bin' && thisObj.binColumn === -1) {
                return;
            }
            if (thisObj.dataPaneColumns.length !== 0 && thisObj.dataPaneColumns.length !== thisObj.dataView.metadata.columns.length) {
                let newIndex: number = value;
                // After update, selected column is undefined or not at same position
                if ((thisObj.dataView.metadata.columns[value] === undefined) ||
                    thisObj.dataPaneColumns[value].displayName !==
                    thisObj.dataView.metadata.columns[value].displayName) {
                    let flag: number = 0;
                    // Find its new position
                    for (let iterator: number = 0; iterator < thisObj.dataView.metadata.columns.length; iterator++) {
                        // Find based on its display name
                        if (thisObj.dataPaneColumns[value].displayName ===
                            thisObj.dataView.metadata.columns[iterator].displayName) {
                            // Found the column. Assign its new index to local variable
                            flag = 1;
                            newIndex = iterator;
                            break;
                        }
                    }

                    if (flag === 0) {
                        // If that column is not found then assign first as column
                        newIndex = 0;
                    }
                }

                if (column === 'bin') {
                    thisObj.binColumn = newIndex;
                } else {
                    if (column === 'target') {
                        thisObj.targetColumn = newIndex;
                    } else {
                        if (column === 'group') {
                            thisObj.groupedColumn = newIndex;
                        } else {
                            thisObj.colorByColumn = newIndex;
                        }
                    }
                }
                thisObj.persistOrient();
            }
        }

        // tslint:disable-next-line:cyclomatic-complexity
        public update(options: VisualUpdateOptions): void {
            const thisObj: this = this;
            thisObj.isCategory = true;
            thisObj.renderedTime++;
            thisObj.viewModel = visualTransform(options, thisObj.host);
            thisObj.width = options.viewport.width;
            thisObj.height = options.viewport.height;
            thisObj.dataView = options.dataViews[0];
            this.currentViewport = {
                height: Math.max(0, options.viewport.height),
                width: Math.max(0, options.viewport.width)
            };

            thisObj.renderLegend(options.dataViews[0]);

            const defaultColor: Fill = {
                solid: {
                    color: '#000000'
                }
            };
            if (options.type === 4 || options.type === 2 || options.type === 36) {
                thisObj.updateCalled = true;
                thisObj.type = true;
            }
            thisObj.settings = Visual.parseSettings(options && options.dataViews && options.dataViews[0]);

            if (thisObj.settings.fontSettings.fontSize > 25) {
                thisObj.settings.fontSettings.fontSize = 25;
            }
            // Limit of range of display ratio
            if (thisObj.settings.ratio.percent < 8) {
                thisObj.settings.ratio.percent = 8;
            }

            if (thisObj.settings.ratio.percent > 80) {
                thisObj.settings.ratio.percent = 80;
            }
            this.currentDataview = options.dataViews[0];
            const textProperties: TextProperties = {
                text: '$123.55',
                fontFamily: thisObj.settings.fontSettings.fontFamily,
                fontSize: `${thisObj.settings.fontSettings.fontSize}px`
            };
            thisObj.textSize = TextMeasurementService.measureSvgTextHeight(textProperties);
            thisObj.textWidth = TextMeasurementService.measureSvgTextWidth(textProperties);
            d3.selectAll('.topCont').remove();
            thisObj.maxLineIsOn = this.settings.analytics.maxLine;
            thisObj.minLineIsOn = this.settings.analytics.minLine;
            thisObj.avgLineIsOn = this.settings.analytics.avgLine;
            thisObj.maxLineStyle = this.settings.analytics.maxLineStyle;
            thisObj.minLineStyle = this.settings.analytics.minLineStyle;
            thisObj.avgLineStyle = this.settings.analytics.avgLineStyle;
            thisObj.showXAxisIsOn = this.settings.xaxis.show;
            thisObj.showLegend = this.settings.legendLabel.show;
            thisObj.maxLineFill = this.settings.analytics.maxLineFill;
            thisObj.minLineFill = this.settings.analytics.minLineFill;
            thisObj.avgLineFill = this.settings.analytics.avgLineFill;
            thisObj.maxLineOpacity = this.settings.analytics.maxLineOpacity;
            thisObj.minLineOpacity = this.settings.analytics.minLineOpacity;
            thisObj.avgLineOpacity = this.settings.analytics.avgLineOpacity;
            thisObj.maxLineDataLabel = this.settings.analytics.maxLineDataLabel;
            thisObj.minLineDataLabel = this.settings.analytics.minLineDataLabel;
            thisObj.avgLineDataLabel = this.settings.analytics.avgLineDataLabel;
            thisObj.constantLineIsOn = this.settings.analytics.constantLine;
            thisObj.constantLineValue = this.settings.analytics.constantLineValue;
            thisObj.constantLineStyle = this.settings.analytics.constantLineStyle;
            thisObj.constantLineFill = this.settings.analytics.constantLineFill;
            thisObj.constantLineOpacity = this.settings.analytics.constantLineOpacity;
            thisObj.constantLineDataLabel = this.settings.analytics.constantLineDataLabel;
            thisObj.value = this.settings.value.displayValue;

            // To manage removal of columns from data pane
            thisObj.setColumns('bin', thisObj.binColumn);
            thisObj.setColumns('target', thisObj.targetColumn);
            thisObj.setColumns('group', thisObj.groupedColumn);
            thisObj.setColumns('colorBy', thisObj.colorByColumn);
            thisObj.dataPaneColumns = options.dataViews[0].metadata.columns;

            if (null === thisObj.groupedColumn || undefined === thisObj.groupedColumn) {
                $('.chartTypeMenu').hide();
                thisObj.groupedColumn = 0;
            }
            if (null === thisObj.binColumn || undefined === thisObj.binColumn) {
                thisObj.binColumn = 1;
            }
            if (null === thisObj.colorByColumn || undefined === thisObj.colorByColumn) {
                thisObj.colorByColumn = 1;
            }
            if (null === thisObj.targetColumn || undefined === thisObj.targetColumn) {
                thisObj.targetColumn = 1;
            }
            if (null === thisObj.chartType || undefined === thisObj.chartType) {
                thisObj.chartType = 'Bar';
            }

            thisObj.setOrient(options);
            this.renderChart();
            this.setMargin();
            if (thisObj.width < $('.chart').width() || thisObj.width < $('.brickChart').width()
                || thisObj.width < $('.columnChart').width()) {
                d3.selectAll('.chart, .brickChart, columnChart').style('margin-left', '10px');
            }

            d3.selectAll('.mainCont, .chart, .columChart, .brickChart').on('click', function (): void {
                thisObj.hideMenus('none');
            });
            thisObj.type = false;
            if (thisObj.selectionIndexes.length > 0) {
                thisObj.$xAxisLabel.hide();
                thisObj.$yAxisLabel.hide();
                $('.menuX').hide();
                $('.menuY').hide();
                $('.menuBinningby').hide();
                thisObj.$xAxis.hide();
                thisObj.$yAxis.hide();
                thisObj.$binningCont.hide();

                d3.select('.colorCont.shape').style('left', $('.menuColor').position().left + thisObj.px);
                d3.select('.colorCont.text').style('left', $('.menuTextColor').position().left + thisObj.px);
                d3.select('.colorByCont').style('left', $('.menuColorBy').position().left + thisObj.px);
            }
            d3.select('.presentationCont.xAxis').style('margin-left', function (): string {
                let parentWidth: number;
                parentWidth = $(this).width();

                return `${(thisObj.width / 2) - (parentWidth / 2) - 30}px`;
            });
        }
        // tslint:disable-next-line:cyclomatic-complexity
        private setMargin(): void {
            let thisObj: this;
            thisObj = this;
            if (!thisObj.settings.presentation.show) {
                let widthofTop10Elements: number = 0;
                let widthofTop9Elements: number = 0;
                let widthofTop6Elements: number = 0;
                let widthofTop4Elements: number = 0;
                let widthofTop3Elements: number = 0;
                let widthofTop2Elements: number = 0;
                let widthofTop1Elements: number = 0;
                const marginbottom: number = 70;
                const widthofTop7LiElements: number = 770;
                const widthofTop6LiElements: number = 660;
                const widthofTop3LiElements: number = 345;
                const widthofTop4LiElements: number = 448;

                const topMenuOptions: string[] = ['menuViewAs', 'menuBinningby', 'menuX', 'menuY',
                    'menuColor', 'menuTextColor', 'menuRangeforBinning', 'menuIsolate', 'menuReset', 'menuUndo'];
                // tslint:disable-next-line:no-any
                let widthArray: any;
                widthArray = [];
                let kCounter: number = 0;
                let iCounter: number;
                for (iCounter = 0; iCounter < topMenuOptions.length; iCounter++) {
                    if ($(thisObj.dot + topMenuOptions[iCounter]).is(':visible')) {
                        widthArray[kCounter] = Math.floor($(thisObj.dot + topMenuOptions[iCounter]).width() + 10);
                        if (kCounter < 10) {
                            widthofTop10Elements = widthofTop10Elements + widthArray[kCounter];
                        }
                        if (kCounter < 9) {
                            widthofTop9Elements = widthofTop9Elements + widthArray[kCounter];
                        }
                        if (kCounter < 6) {
                            widthofTop6Elements = widthofTop6Elements + widthArray[kCounter];
                        }
                        if (kCounter < 4) {
                            widthofTop4Elements = widthofTop4Elements + widthArray[kCounter];
                        }
                        if (kCounter < 3) {
                            widthofTop3Elements = widthofTop3Elements + widthArray[kCounter];
                        }
                        if (kCounter < 2) {
                            widthofTop2Elements = widthofTop2Elements + widthArray[kCounter];
                        }
                        if (kCounter < 1) {
                            widthofTop1Elements = widthofTop1Elements + widthArray[kCounter];
                        }
                        kCounter++;
                    }
                }
                d3.select('.mainCont').style({
                    height: thisObj.width >= widthofTop10Elements
                        ? thisObj.height - 160 + thisObj.px : thisObj.width >= widthofTop4Elements
                            ? thisObj.height - 200 + thisObj.px : thisObj.width >= widthofTop3Elements
                                ? thisObj.height - 216 + thisObj.px : thisObj.width >= widthofTop2Elements
                                    ? thisObj.height - 289 + thisObj.px : thisObj.height - (widthArray.length * 40) + thisObj.px
                });
                d3.select('.mainCont').style({
                    'margin-top': thisObj.width >= widthofTop9Elements
                        ? 28 + thisObj.px : thisObj.width >= widthofTop4Elements
                            ? 78 + thisObj.px : thisObj.width >= widthofTop3Elements
                                ? 130 + thisObj.px : 148 + thisObj.px
                });
                if (thisObj.width <= 1000) {
                    d3.select('.topCont').style({
                        // tslint:disable-next-line:prefer-template
                        'margin-bottom': marginbottom + thisObj.px
                    });
                }

                if (thisObj.width <= widthofTop4LiElements && thisObj.width >= widthofTop3LiElements) {
                    d3.select('.mainCont').style({
                        // tslint:disable-next-line:prefer-template
                        'margin-top': '100' + thisObj.px
                    });
                }
            }
        }
        // tslint:disable-next-line:no-any
        private hideMenus(menu: any): void {
            const thisObj: this = this;
            if (thisObj.flag === 1) {
                thisObj.flag = 0;
                thisObj.selectedColumnsString = thisObj.dataColIndex.join('-');
                thisObj.persistOrient();
                thisObj.setOrient(thisObj.options);
            }
            menu === thisObj.$binningCont ? thisObj.$binningCont.slideToggle() : thisObj.$binningCont.hide();
            menu === thisObj.$colorByCont ? thisObj.$colorByCont.slideToggle() : thisObj.$colorByCont.hide();
            menu === thisObj.$groupingCont ? thisObj.$groupingCont.slideToggle() : thisObj.$groupingCont.hide();
            menu === thisObj.$targetCont ? thisObj.$targetCont.slideToggle() : thisObj.$targetCont.hide();
            menu === thisObj.$chartTypeMenu ? thisObj.$chartTypeMenu.slideToggle() : thisObj.$chartTypeMenu.hide();
            menu === thisObj.$xAxis ? thisObj.$xAxis.slideToggle() : thisObj.$xAxis.hide();
            menu === thisObj.$yAxis ? thisObj.$yAxis.toggle('slide') : thisObj.$yAxis.hide();
            menu === thisObj.$colorContText ? thisObj.$colorContText.slideToggle() : thisObj.$colorContText.hide();
            menu === thisObj.$colorContShape ? thisObj.$colorContShape.slideToggle() : thisObj.$colorContShape.hide();
            menu === thisObj.$dynamicBinCont ? thisObj.$dynamicBinCont.slideToggle() : thisObj.$dynamicBinCont.hide();
            if (thisObj.$columnSelector !== undefined) {
                menu === thisObj.$columnSelector ? thisObj.$columnSelector.slideToggle() : thisObj.$columnSelector.hide();
            }
        }
        // tslint:disable-next-line:no-any
        private getTooltipDataAnalyticsLine(val: any, data: any, type: string, display: string, text: number, thisObj: this)
            : VisualTooltipDataItem[] {
            // tslint:disable-next-line:prefer-const
            let displayName: string;
            const lineToolTipValue: VisualTooltipDataItem[] = [];
            // tslint:disable-next-line:prefer-const
            let value: string;
            if (type === 'line') {

                lineToolTipValue.push({
                    displayName: 'Max',
                    value: thisObj.maxValue === 0 ? thisObj.maxValue.toString() :
                        thisObj.targetColumnformatter.format(thisObj.maxValue).toString()
                });
                lineToolTipValue.push({
                    displayName: 'Min',
                    value: thisObj.minValue === 0 ? thisObj.minValue.toString() :
                        thisObj.targetColumnformatter.format(thisObj.minValue).toString()
                });
                lineToolTipValue.push({
                    displayName: 'Avg',
                    value: thisObj.averageValue === 0 ? thisObj.averageValue.toString() :
                        thisObj.targetColumnformatter.format(thisObj.averageValue).toString()
                });
                lineToolTipValue.push({
                    displayName: 'Constant',
                    value: thisObj.constantLineValue === 0 ? thisObj.constantLineValue.toString() :
                        thisObj.targetColumnformatter.format(thisObj.constantLineValue).toString()
                });

                return lineToolTipValue;
            }
        }
        // tslint:disable-next-line:no-any
        // tslint:disable-next-line:cyclomatic-complexity
        // tslint:disable-next-line:no-any
        private getTooltipData(val: any, data: any, type: string, display: string, text: number, thisObj: this):
            VisualTooltipDataItem[] {
            let displayName: string;
            const lineToolTipValue: VisualTooltipDataItem[] = [];
            let value: string;
            if (type === 'inDiv' || type === 'bar') {
                if (data !== undefined && data[`key`] !== undefined) {
                    if (type === 'inDiv') {
                        displayName = data[`key`].toString();
                        value = undefined === data[`value`] ? '' : thisObj.targetColumnformatter.format(data[`value`]).toString();
                    } else {
                        displayName = data[`key`].toString();
                        // tslint:disable-next-line:max-line-length
                        value = (undefined === data[`values`].value) ? '' : thisObj.targetColumnformatter.format(data[`values`].value).toString();
                    }
                }
            } else if (type === 'brick') {
                // tslint:disable-next-line:typedef
                const arrColorBar = this.barcolor;
                // tslint:disable-next-line:no-any
                let inDivColor: any;
                // tslint:disable-next-line:typedef
                for (let valueIndex = 0; valueIndex < val.values[0].length; valueIndex++) {
                    // tslint:disable-next-line:typedef
                    for (let colorSelection = 0; colorSelection < arrColorBar.length; colorSelection++) {
                        // tslint:disable-next-line:max-line-length
                        if (arrColorBar[colorSelection][`key`].toString() === val.values[0][valueIndex][`key`].toString() && arrColorBar[colorSelection][`key`].toString() !== null) {
                            // tslint:disable-next-line:no-unused-expression
                            inDivColor = arrColorBar[colorSelection][`value`].toString();
                            break;
                        }
                    }
                    lineToolTipValue.push({
                        // tslint:disable-next-line:max-line-length
                        displayName: undefined === val.values[0][valueIndex][`key`].toString() ? '' : thisObj.targetColumnformatter.format(val.values[0][valueIndex][`key`]).toString(),
                        // tslint:disable-next-line:max-line-length
                        value: val.values[0][valueIndex][`values`][`value`] % 1 === 0 ?
                            val.values[0][valueIndex][`values`][`value`].toString() :
                                                                        val.values[0][valueIndex][`values`][`value`].toFixed(2),
                        color: inDivColor
                    });
                }

                return lineToolTipValue;
            } else {
                if (data !== undefined && data[`key`] !== undefined) {
                    displayName = data[`key`].toString();
                    value = undefined === data[`value`] ? '' : thisObj.targetColumnformatter.format(data[`value`]).toString();
                }
            }
            if (undefined !== displayName && undefined !== value) {
                return [{
                    displayName: displayName.toString(),
                    value: value.toString()
                }];
            }
        }

        private static parseSettings(dataView: DataView): VisualSettings {
            return VisualSettings.parse(dataView) as VisualSettings;
        }
        // tslint:disable-next-line:no-any
        private renderLegend(dataView: any): void {
            // tslint:disable-next-line:typedef
            const a = [{
                label: 'ABC',
                color: 'black',
                icon: powerbi.extensibility.utils.chart.legend.LegendIcon.Box,
                selected: true,
                identity: this.host.createSelectionIdBuilder().withCategory(
                    dataView.categorical.categories[0], 0).createSelectionId()
            }];
        }

        /**
         * This function gets called for each of the objects defined in the capabilities files and allows us to select which of the
         * objects and properties we want to expose to the users in the property pane.
         *
         */
        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions):
            VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
            const objectEnumeration: VisualObjectInstance[] = [];
            if (options.objectName === 'presentation') {
                objectEnumeration.push({
                    objectName: options.objectName,
                    properties:
                        {
                            show: this.settings.presentation.show
                        },
                    selector: null
                });
            } else if (options.objectName === 'legendLabel' && this.colorByColumn !== -1 && this.chartType.toLowerCase() !== 'table') {
                objectEnumeration.push({
                    objectName: options.objectName,
                    properties:
                        {
                            show: this.settings.legendLabel.show
                        },
                    selector: null
                });
            } else if (options.objectName === 'xaxis' && this.chartType.toLowerCase() === 'column') {
                objectEnumeration.push({
                    objectName: options.objectName,
                    properties:
                        {
                            show: this.settings.xaxis.show
                        },
                    selector: null
                });
            } else if (options.objectName === 'analytics') {
                if (this.chartType.toLowerCase() === 'bar' || this.chartType.toLowerCase() === 'column') {
                    // tslint:disable-next-line:no-any
                    const objProps: any = {};
                    objProps[`maxLine`] = this.settings.analytics.maxLine;
                    if (this.settings.analytics.maxLine) {
                        objProps[`maxLineStyle`] = this.settings.analytics.maxLineStyle;
                        objProps[`maxLineFill`] = this.settings.analytics.maxLineFill;
                        objProps[`maxLineOpacity`] = this.settings.analytics.maxLineOpacity;
                        objProps[`maxLineDataLabel`] = this.settings.analytics.maxLineDataLabel;
                    }

                    objProps[`minLine`] = this.settings.analytics.minLine;
                    if (this.settings.analytics.minLine) {
                        objProps[`minLineStyle`] = this.settings.analytics.minLineStyle;
                        objProps[`minLineFill`] = this.settings.analytics.minLineFill;
                        objProps[`minLineOpacity`] = this.settings.analytics.minLineOpacity;
                        objProps[`minLineDataLabel`] = this.settings.analytics.minLineDataLabel;
                    }
                    objProps[`avgLine`] = this.settings.analytics.avgLine;
                    if (this.settings.analytics.avgLine) {
                        objProps[`avgLineStyle`] = this.settings.analytics.avgLineStyle;
                        objProps[`avgLineFill`] = this.settings.analytics.avgLineFill;
                        objProps[`avgLineOpacity`] = this.settings.analytics.avgLineOpacity;
                        objProps[`avgLineDataLabel`] = this.settings.analytics.avgLineDataLabel;
                    }
                    objProps[`constantLine`] = this.settings.analytics.constantLine;
                    if (this.settings.analytics.constantLine) {
                        objProps[`constantLineValue`] = this.settings.analytics.constantLineValue;
                        objProps[`constantLineStyle`] = this.settings.analytics.constantLineStyle;
                        objProps[`constantLineFill`] = this.settings.analytics.constantLineFill;
                        objProps[`constantLineOpacity`] = this.settings.analytics.constantLineOpacity;
                        objProps[`constantLineDataLabel`] = this.settings.analytics.constantLineDataLabel;
                    }
                    objectEnumeration.push({
                        objectName: options.objectName,
                        properties: objProps,
                        selector: null
                    });
                }
            } else if (options.objectName === 'ratio') {
                if (this.chartType.toLowerCase() === 'bar') {
                    objectEnumeration.push({
                        objectName: options.objectName,
                        properties:
                            {
                                percent: this.settings.ratio.percent
                            },
                        selector: null
                    });
                }
            } else if (options.objectName === 'value') {
                if (this.chartType.toLowerCase() === 'bar') {
                    objectEnumeration.push({
                        objectName: options.objectName,
                        properties:
                            {
                                displayValue: this.settings.value.displayValue

                            },
                        selector: null
                    });
                }
            } else if (options.objectName === 'fontSettings') {
                objectEnumeration.push({
                    objectName: options.objectName,
                    properties:
                        {
                            fontSize: this.settings.fontSettings.fontSize,
                            fontFamily: this.settings.fontSettings.fontFamily
                        },
                    selector: null
                });
            }

            return objectEnumeration;
        }
    }
}
