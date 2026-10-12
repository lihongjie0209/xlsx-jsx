import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook theme="excel2023">
    <Worksheet name="Summary">
      <Row><Cell>Product</Cell><Cell>Revenue</Cell></Row>
      <Row><Cell>Widget</Cell><Cell value={1250} /></Row>
    </Worksheet>
  </Workbook>
);
