import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Summary" autoFit={true}>
    <Row><Cell>Product</Cell><Cell>Quantity</Cell><Cell>Comment</Cell></Row>
    <Row><Cell>Apples</Cell><Cell value={12} /><Cell>Arrived this morning</Cell></Row>
    <Row><Cell>Blueberries</Cell><Cell value={240} /><Cell>Requires refrigerated storage</Cell></Row>
  </Worksheet>
</Workbook>;
