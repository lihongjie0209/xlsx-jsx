import { Workbook, Worksheet, Column, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Columns">
    <Column min={1} max={1} widthPixels={140} />
    <Column min={2} max={3} widthPixels={160} />
    <Column min={4} max={4} hidden={true} />
    <Row>
      <Cell>First</Cell>
      <Cell>Grouped width</Cell>
      <Cell>Grouped width</Cell>
      <Cell>Hidden</Cell>
    </Row>
  </Worksheet>
</Workbook>;
