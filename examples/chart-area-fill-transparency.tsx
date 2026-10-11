import { Workbook, Worksheet, Chart, Series, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <Chart
      type="column"
      cell="D2"
      chartAreaFillColor="#FFFFFF"
      chartAreaFillTransparency={25}
      plotAreaFillColor="#D9EAF7"
      plotAreaFillTransparency={40}
    >
      <Series values="Data!B2:B3" categories="Data!A2:A3" />
    </Chart>
    <Row><Cell>Quarter</Cell><Cell>Revenue</Cell></Row>
    <Row><Cell>Q1</Cell><Cell value={10} /></Row>
    <Row><Cell>Q2</Cell><Cell value={20} /></Row>
  </Worksheet>
</Workbook>;
