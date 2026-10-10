import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Font emphasis">
    <Row>
      <Cell bold={true}>Bold text</Cell>
      <Cell italic={true} value={42} />
      <Cell bold={true} formula="=B1+1" />
    </Row>
  </Worksheet>
</Workbook>;
