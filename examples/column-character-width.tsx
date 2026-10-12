import { Workbook, Worksheet, Column, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Column widths">
    <Column min={1} max={1} width={8.43} />
    <Column min={2} max={2} width={12.5} />
    <Row><Cell>Item</Cell><Cell>Quantity</Cell></Row>
    <Row><Cell>Notebook</Cell><Cell value={4} /></Row>
  </Worksheet>
</Workbook>;
