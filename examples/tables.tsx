import { Workbook, Worksheet, Table, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Sales">
    <Table range="A1:B3" name="SalesTable" style="TableStyleMedium4" />
    <Row><Cell>Product</Cell><Cell>Amount</Cell></Row>
    <Row><Cell>Tea</Cell><Cell value={12} /></Row>
    <Row><Cell>Coffee</Cell><Cell value={18} /></Row>
  </Worksheet>
</Workbook>;
