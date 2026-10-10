import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Font Decorations">
    <Row>
      <Cell underline="single">Single underline</Cell>
      <Cell underline="double">Double underline</Cell>
      <Cell underline="singleAccounting">Accounting underline</Cell>
      <Cell underline="doubleAccounting" strikethrough={true}>Strike-through</Cell>
    </Row>
  </Worksheet>
</Workbook>;
