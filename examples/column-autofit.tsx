import { Workbook, Worksheet, Column, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Column behavior">
    <Column min={2} max={2} autoFitWidthPixels={144} />
    <Row><Cell>Product</Cell><Cell>Quantity</Cell></Row>
    <Row><Cell>Apples</Cell><Cell value={12} /></Row>
  </Worksheet>
</Workbook>;
