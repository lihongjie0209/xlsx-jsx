import { Workbook, Worksheet, Row, Cell, Chart, Series } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <Row><Cell>Month</Cell><Cell>Revenue</Cell><Cell>Expenses</Cell></Row>
    <Row><Cell>Jan</Cell><Cell value={12} /><Cell value={7} /></Row>
    <Row><Cell>Feb</Cell><Cell value={18} /><Cell value={9} /></Row>
    <Row><Cell>Mar</Cell><Cell value={15} /><Cell value={6} /></Row>
  </Worksheet>
  <Worksheet name="Sales">
    <Chart type="column" cell="D2" width={480} height={288} name="Revenue and expenses" title="Monthly revenue &amp; expenses — 2026" xAxisTitle="Month" yAxisTitle="Revenue ($)" altText="Monthly revenue compared with expenses">
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" />
      <Series values="Data!$C$2:$C$4" categories="Data!$A$2:$A$4" title="Expenses &amp; travel" />
    </Chart>
    <Chart type="line" cell="K2" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" />
    </Chart>
    <Chart type="bar" cell="D20" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" />
    </Chart>
    <Chart type="pie" cell="K20" width={480} height={288}>
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" name="Data!$B$1" />
    </Chart>
  </Worksheet>
</Workbook>;
