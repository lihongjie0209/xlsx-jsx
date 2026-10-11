import { Workbook, Worksheet, Row, Cell, Chart, Series, ErrorBars, Trendline, ChartDataTable } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <Row><Cell>Month</Cell><Cell>Revenue</Cell><Cell>Expenses</Cell><Cell>Plus error</Cell><Cell>Minus error</Cell></Row>
    <Row><Cell>Jan</Cell><Cell value={12} /><Cell value={7} /><Cell value={1.5} /><Cell value={0.8} /></Row>
    <Row><Cell>Feb</Cell><Cell value={18} /><Cell value={9} /><Cell value={1.2} /><Cell value={0.9} /></Row>
    <Row><Cell>Mar</Cell><Cell value={15} /><Cell value={6} /><Cell value={1.8} /><Cell value={0.6} /></Row>
  </Worksheet>
  <Worksheet name="Sales">
    <Chart type="column" cell="D2" width={480} height={288} name="Revenue and expenses" title="Monthly revenue &amp; expenses — 2026" titleOverlay={true} chartAreaFillColor="#FFFFFF" plotAreaFillColor="#F2F2F2" xAxisTitle="Month" yAxisTitle="Revenue ($)" yAxisNumberFormat="$#,##0.00" yAxisMin={0} yAxisMax={20} yAxisMajorUnit={5} yAxisMinorUnit={1} valueAxisDisplayUnit="thousands" categoryAxisLabelAlignment="left" categoryAxisOnTickMarks={true} style={7} xAxisCrossing="max" yAxisCrossing="min" horizontalAxisHidden={true} horizontalAxisMinorGridlines={true} verticalAxisMajorGridlines={{color: "#D9E1F2", width: 0.75, dash: "solid"}} horizontalAxisMajorTickType="none" verticalAxisMajorTickType="outside" horizontalAxisLabelPosition="high" verticalAxisLabelPosition="low" legendPosition="bottom" legendOverlay={true} showHiddenData={true} emptyCellsAs="span" showNaAsEmptyCell={true} altText="Monthly revenue compared with expenses">
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" gap={70} overlap={-25} showValue={true} showCategoryName={true} labelPosition="insideEnd" labelNumberFormat="$#,##0.00" labelFontSize={12} labelFontBold={true} labelFontColor="#336699">
        <ErrorBars axis="y" type="custom" plusRange="Data!$D$2:$D$4" minusRange="Data!$E$2:$E$4" lineColor="#4472C4" lineWidth={1.5} lineDash="dash" lineTransparency={20} />
        <Trendline type="linear" deleteFromLegend={true} />
      </Series>
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Expenses &amp; travel" deleteFromLegend={true} />
      <ChartDataTable showHorizontalBorders={false} showVerticalBorders={true} showOutlineBorders={true} showLegendKeys={true} />
    </Chart>
    <Chart type="line" cell="K2" width={480} height={288} showLegend={false} valueAxisReversed={true} categoryAxisLabelInterval={2} categoryAxisTickInterval={2} dropLines={true}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" markerType="circle" markerSize={8} lineColor="#336699" lineWidth={1.5} lineDash="dashDot" smooth={true} showValue={true} showLegendKey={true} showLeaderLines={true} labelSeparator="semicolon">
        <Trendline type="linear" forward={2} displayEquation={true} displayRSquared={true} />
      </Series>
    </Chart>
    <Chart type="bar" cell="D20" width={480} height={288} verticalAxisHidden={true} autoTitleDeleted={true}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" invertIfNegative={true}>
        <ErrorBars axis="x" type="standardError" />
      </Series>
    </Chart>
    <Chart type="pie" cell="K20" width={480} height={288} firstSliceAngle={210} legendDeletedEntries={[1]}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" pointColors={{0: "#4472C4", 2: "#ED7D31"}} showPercentage={true} />
    </Chart>
    <Chart type="area" cell="D38" width={480} height={288} dropLines={true}>
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
    <Chart type="scatterWithLines" cell="K146" width={480} height={288} xAxisMin={10} xAxisMax={20} yAxisMin={5} yAxisMax={10} xAxisLogBase={10} yAxisLogBase={10} xAxisCrossing="min" yAxisCrossing="max" xAxisTitle="Units" yAxisTitle="Revenue" xAxisReversed={true} yAxisReversed={true} xAxisDisplayUnit="thousands" xAxisDisplayUnitLabel={false} yAxisDisplayUnit="thousands" xAxisMajorUnit={2} xAxisMinorUnit={1} yAxisMajorUnit={2} yAxisMinorUnit={1} xAxisNumberFormat="0.0" yAxisNumberFormat="0.00">
      <Series values="Data!$C$2:$C$4" categories="Data!$B$2:$B$4" title="Revenue scatter line" />
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
    <Chart type="column" cell="K218" title="Manually positioned chart title" titleFontSize={16} titleFontBold={false} titleFontItalic={true} titleFontUnderline={true} titleFontRotation={45} titleFontColor="#336699" titleFontName="Aptos" titleFontPitchFamily={34} titleFontCharset={127} titleLayout={{x: 0, y: 0.15}} plotAreaLayout={{x: 0, y: 0.3, width: 0.7, height: 0.5}} legendLayout={{x: 0.1, y: 0, width: 0.3, height: 0.4}}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Manually positioned plot area" />
    </Chart>
    <Chart type="column" cell="D236" title="Stacked title text" titleFontRotation={270}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Stacked title example" />
    </Chart>
    <Chart type="column" cell="K236" title="East Asian vertical title" titleFontRotation={271}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" title="Vertical title example" />
    </Chart>
  </Worksheet>
  <Worksheet name="Timeline">
    <Chart type="line" cell="D2" title="Monthly revenue" categoryAxisType="date" categoryAxisMajorTimeUnit="months" categoryAxisMinorTimeUnit="days" categoryAxisMinDate="2026-01-01" categoryAxisMaxDate="2026-03-01">
      <Series values="Timeline!$B$2:$B$4" categories="Timeline!$A$2:$A$4" title="Revenue" />
    </Chart>
    <Row><Cell>Date</Cell><Cell>Revenue</Cell></Row>
    <Row><Cell date="2026-01-01" /><Cell value={120} /></Row>
    <Row><Cell date="2026-02-01" /><Cell value={135} /></Row>
    <Row><Cell date="2026-03-01" /><Cell value={142} /></Row>
  </Worksheet>
</Workbook>;
