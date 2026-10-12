import { Workbook, Worksheet, Chartsheet, Chart, Series, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <Row><Cell value="Month" /><Cell value="Sales" /></Row>
    <Row><Cell value="Jan" /><Cell value={10} /></Row>
    <Row><Cell value="Feb" /><Cell value={20} /></Row>
  </Worksheet>
  <Chartsheet name="Sales chart">
    <Chart type="column">
      <Series categories="Data!$A$2:$A$3" values="Data!$B$2:$B$3" title="Sales" />
    </Chart>
  </Chartsheet>
</Workbook>;
