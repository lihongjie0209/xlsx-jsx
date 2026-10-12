import { Workbook, Worksheet, Column, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="CenterAcross">
    <Column min={1} max={2} horizontalAlign="centerAcross" />
    <Row horizontalAlign="centerAcross">
      <Cell>Quarterly report</Cell><Cell />
    </Row>
    <Row>
      <Cell horizontalAlign="centerAcross">Totals</Cell><Cell />
    </Row>
  </Worksheet>
</Workbook>;
