import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Alignment">
    <Row>
      <Cell horizontalAlign="center" verticalAlign="center">Centered</Cell>
      <Cell horizontalAlign="right" verticalAlign="top" value={42} />
      <Cell horizontalAlign="left" formula="=B1+1" />
    </Row>
    <Row>
      <Cell textRotation={-45}>Angled up</Cell>
      <Cell textRotation={45}>Angled down</Cell>
      <Cell textRotation={270}>Stacked vertical</Cell>
    </Row>
    <Row height={40.5}>
      <Cell verticalAlign="justify">Vertically justified</Cell>
      <Cell verticalAlign="distributed">Vertically distributed</Cell>
    </Row>
  </Worksheet>
</Workbook>;
