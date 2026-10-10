import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Dynamic formulas">
    <Row>
      <Cell>Foo</Cell>
      <Cell formula="=LEN(A1:A3)" dynamicArray={true} arrayFormulaRange="B1:B3" />
      <Cell />
      <Cell formula={'=FILTER(A1:A3,A1:A3<>"")'} dynamicArray={true} />
    </Row>
    <Row><Cell>Food</Cell></Row>
    <Row><Cell>Frood</Cell></Row>
  </Worksheet>
</Workbook>;
