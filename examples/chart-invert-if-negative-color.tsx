import { Workbook, Worksheet, Row, Cell, Chart, Series } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Data">
    <Chart type="column" cell="F2" width={480} height={288}>
      <Series values="Data!$A$2:$A$4" invertIfNegativeColor="#FF6666" />
      <Series values="Data!$B$2:$B$4" invertIfNegative={true} />
    </Chart>
    <Row><Cell>North</Cell><Cell>South</Cell></Row>
    <Row><Cell value={30} /><Cell value={22} /></Row>
    <Row><Cell value={24} /><Cell value={-10} /></Row>
    <Row><Cell value={32} /><Cell value={16} /></Row>
  </Worksheet>
</Workbook>;
