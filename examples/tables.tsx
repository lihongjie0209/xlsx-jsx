import { Workbook, Worksheet, Table, TableColumn, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Sales">
    <Table range="A1:B3" name="SalesTable" style="TableStyleMedium4" />
    <Table range="D1:E3" name="RawSales" headerRow={false}>
      <TableColumn name="Item" />
      <TableColumn name="Quantity" />
    </Table>
    <Row><Cell>Product</Cell><Cell>Amount</Cell><Cell /><Cell>Tea</Cell><Cell value={12} /></Row>
    <Row><Cell>Tea</Cell><Cell value={12} /><Cell /><Cell>Coffee</Cell><Cell value={18} /></Row>
    <Row><Cell>Coffee</Cell><Cell value={18} /><Cell /><Cell>Juice</Cell><Cell value={9} /></Row>
  </Worksheet>
</Workbook>;
