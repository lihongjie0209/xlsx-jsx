import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Inventory" autoFilter="A1:B4" autoFilterColumns={[{ column: 0, values: ["Apples", "Oranges"], includeBlanks: true }]}>
    <Row><Cell>Product</Cell><Cell>Quantity</Cell></Row>
    <Row><Cell>Apples</Cell><Cell value={12} /></Row>
    <Row><Cell>Oranges</Cell><Cell value={8} /></Row>
    <Row><Cell>Bananas</Cell><Cell value={5} /></Row>
  </Worksheet>
</Workbook>;
