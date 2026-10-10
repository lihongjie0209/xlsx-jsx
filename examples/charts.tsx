import { Workbook, Worksheet, Row, Cell, Chart, Series } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <Row><Cell>Month</Cell><Cell>Revenue</Cell><Cell>Expenses</Cell></Row>
    <Row><Cell>Jan</Cell><Cell value={12} /><Cell value={7} /></Row>
    <Row><Cell>Feb</Cell><Cell value={18} /><Cell value={9} /></Row>
    <Row><Cell>Mar</Cell><Cell value={15} /><Cell value={6} /></Row>
  </Worksheet>
  <Worksheet name="Sales">
    <Chart type="column" cell="D2" width={480} height={288} name="Revenue and expenses" title="Monthly revenue &amp; expenses — 2026" xAxisTitle="Month" yAxisTitle="Revenue ($)" legendPosition="bottom" altText="Monthly revenue compared with expenses">
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Expenses &amp; travel" />
    </Chart>
    <Chart type="line" cell="K2" width={480} height={288} showLegend={false}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" />
    </Chart>
    <Chart type="bar" cell="D20" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" />
    </Chart>
    <Chart type="pie" cell="K20" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" />
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
  </Worksheet>
</Workbook>;
