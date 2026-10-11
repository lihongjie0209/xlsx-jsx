import { Workbook, Worksheet, Row, Cell, Chart, Series } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <Row><Cell>Category</Cell><Cell>Value</Cell></Row>
    <Row><Cell>One</Cell><Cell value={10} /></Row>
    <Row><Cell>Two</Cell><Cell value={20} /></Row>
    <Row><Cell>Three</Cell><Cell value={30} /></Row>
  </Worksheet>
  <Worksheet name="Chart">
    <Chart type="column" cell="D2">
      <Series values="Data!$B$2:$B$4" categories="Data!$A$2:$A$4" customLabels={["Forecast", null, "Actual"]} />
    </Chart>
  </Worksheet>
</Workbook>;
