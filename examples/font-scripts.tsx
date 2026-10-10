import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Font Scripts">
    <Row>
      <Cell>H</Cell>
      <Cell script="subscript">2</Cell>
      <Cell>O</Cell>
      <Cell>X</Cell>
      <Cell script="superscript">2</Cell>
    </Row>
  </Worksheet>
</Workbook>;
