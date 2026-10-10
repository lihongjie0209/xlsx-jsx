import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Wrapping">
    <Row>
      <Cell wrapText={true} verticalAlign="top">
        Longer text wraps inside its cell instead of extending across adjacent columns.
      </Cell>
      <Cell>Adjacent value</Cell>
    </Row>
  </Worksheet>
</Workbook>;
