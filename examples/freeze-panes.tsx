import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet
    name="Monthly report"
    freezeRows={1}
    freezeColumns={1}
    freezeTopLeftCell="C3"
  >
    <Row><Cell>Product</Cell><Cell>January</Cell><Cell>February</Cell></Row>
    <Row><Cell>Apples</Cell><Cell value={12} /><Cell value={9} /></Row>
    <Row><Cell>Oranges</Cell><Cell value={8} /><Cell value={11} /></Row>
  </Worksheet>
</Workbook>;
