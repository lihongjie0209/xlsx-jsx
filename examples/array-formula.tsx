import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Array formula">
    <Row>
      <Cell value={2} />
      <Cell formula="=A1:A2*2" arrayFormulaRange="B1:B2" />
    </Row>
    <Row><Cell value={3} /></Row>
  </Worksheet>
</Workbook>;
