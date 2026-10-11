import { Workbook, Worksheet, Row, Cell, Chart, Series, ErrorBars, ChartDataTable } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <Row><Cell>Month</Cell><Cell>Revenue</Cell><Cell>Expenses</Cell></Row>
    <Row><Cell>Jan</Cell><Cell value={12} /><Cell value={7} /></Row>
    <Row><Cell>Feb</Cell><Cell value={18} /><Cell value={9} /></Row>
    <Row><Cell>Mar</Cell><Cell value={15} /><Cell value={6} /></Row>
  </Worksheet>
  <Worksheet name="Sales">
    <Chart type="column" cell="D2" width={480} height={288} name="Revenue and expenses" title="Monthly revenue &amp; expenses — 2026" xAxisTitle="Month" yAxisTitle="Revenue ($)" yAxisNumberFormat="$#,##0.00" yAxisMin={0} yAxisMax={20} yAxisMajorUnit={5} yAxisMinorUnit={1} categoryAxisLabelAlignment="left" categoryAxisOnTickMarks={true} style={7} xAxisCrossing="max" yAxisCrossing="min" horizontalAxisHidden={true} horizontalAxisMinorGridlines={true} verticalAxisMajorGridlines={false} horizontalAxisMajorTickType="none" verticalAxisMajorTickType="outside" horizontalAxisLabelPosition="high" verticalAxisLabelPosition="low" legendPosition="bottom" altText="Monthly revenue compared with expenses">
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" showValue={true} showCategoryName={true} labelPosition="insideEnd" labelNumberFormat="$#,##0.00">
        <ErrorBars axis="y" type="standardError" />
      </Series>
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Expenses &amp; travel" />
      <ChartDataTable showHorizontalBorders={false} showVerticalBorders={true} showOutlineBorders={true} showLegendKeys={true} />
    </Chart>
    <Chart type="line" cell="K2" width={480} height={288} showLegend={false} valueAxisReversed={true} categoryAxisLabelInterval={2} categoryAxisTickInterval={2}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" markerType="circle" markerSize={8} showValue={true} showLegendKey={true} showLeaderLines={true} labelSeparator="semicolon" />
    </Chart>
    <Chart type="bar" cell="D20" width={480} height={288} verticalAxisHidden={true}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1">
        <ErrorBars axis="x" type="standardError" />
      </Series>
    </Chart>
    <Chart type="pie" cell="K20" width={480} height={288} firstSliceAngle={210}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" showPercentage={true} />
    </Chart>
    <Chart type="area" cell="D38" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Revenue area" />
    </Chart>
    <Chart type="areaStacked" cell="K38" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Stacked revenue" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Stacked expenses" />
    </Chart>
    <Chart type="areaPercentStacked" cell="D56" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Revenue share" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Expense share" />
    </Chart>
    <Chart type="columnStacked" cell="K56" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Stacked columns" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Stacked column details" />
    </Chart>
    <Chart type="columnPercentStacked" cell="D74" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Column share" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Column expense share" />
    </Chart>
    <Chart type="barStacked" cell="K74" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Stacked bars" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Stacked bar details" />
    </Chart>
    <Chart type="barPercentStacked" cell="D92" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Bar share" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Bar expense share" />
    </Chart>
    <Chart type="lineStacked" cell="K92" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Stacked lines" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Stacked line details" />
    </Chart>
    <Chart type="linePercentStacked" cell="D110" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Line share" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Line expense share" />
    </Chart>
    <Chart type="doughnut" cell="K110" width={480} height={288} holeSize={65} firstSliceAngle={45}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Revenue mix" />
    </Chart>
    <Chart type="radar" cell="D128" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Revenue radar" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Expenses radar" />
    </Chart>
    <Chart type="radarWithMarkers" cell="K128" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Revenue radar points" />
    </Chart>
    <Chart type="radarFilled" cell="D146" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Revenue radar filled" />
    </Chart>
    <Chart type="scatterWithLines" cell="K146" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Revenue scatter line" />
    </Chart>
    <Chart type="scatterWithLinesAndMarkers" cell="D164" width={480} height={288}>
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Expense scatter points" />
    </Chart>
    <Chart type="scatterWithSmoothLines" cell="K164" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Revenue smoothed line" />
    </Chart>
    <Chart type="scatterWithSmoothLinesAndMarkers" cell="D182" width={480} height={288}>
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Expense smoothed points" />
    </Chart>
    <Chart type="scatter" cell="K182" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Revenue scatter" />
    </Chart>
    <Chart type="stock" cell="D200" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="High" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Low" />
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Close" />
    </Chart>
    <Chart type="stock" cell="K200" width={480} height={288} xAxisTitle="Date" yAxisTitle="Price">
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Open" />
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="High" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Low" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Close" />
    </Chart>
    <Chart type="column" cell="D218" width={480} height={288} yAxisLogBase={10} yAxisMin={1}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Revenue (log scale)" />
    </Chart>
  </Worksheet>
</Workbook>;
