import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Indentation">
    <Row>
      <Cell horizontalAlign="left" indent={2}>Indented left</Cell>
      <Cell horizontalAlign="distributed" indent={3}>Distributed and indented</Cell>
    </Row>
  </Worksheet>
</Workbook>;
