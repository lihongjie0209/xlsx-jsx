import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Inventory" topLeftCell="H6" selection={{activeCell: "D3", range: "D3:F5"}}>
    <Row><Cell>Item</Cell><Cell>Quantity</Cell></Row>
    <Row><Cell>Paper</Cell><Cell value={240} /></Row>
  </Worksheet>
</Workbook>;
