import { Workbook, Worksheet, Row, Cell, Chart, Series } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <Row><Cell>Point</Cell><Cell>Value</Cell></Row>
    <Row><Cell>A</Cell><Cell value={1} /></Row>
    <Row><Cell>B</Cell><Cell value={1} /></Row>
  </Worksheet>
  <Worksheet name="Dashboard">
    <Chart type="line" cell="B2" decorative={true}>
      <Series values="Data!B2:B3" categories="Data!A2:A3" />
    </Chart>
  </Worksheet>
</Workbook>;
