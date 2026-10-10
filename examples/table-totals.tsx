import { Workbook, Worksheet, Table, TableColumn, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Sales">
    <Table range="A1:C3" name="SalesTable" totalsRow={true}>
      <TableColumn totalLabel="Total" headerFormat={{ bold: true }} />
      <TableColumn totalFunction="sum" />
      <TableColumn formula="[[#This Row],Amount]*0.1" />
    </Table>
    <Row><Cell>Product</Cell><Cell>Amount</Cell><Cell>Tax</Cell></Row>
    <Row><Cell>Tea</Cell><Cell value={12} /><Cell value={1.2} /></Row>
    <Row><Cell>Total</Cell><Cell /><Cell /></Row>
  </Worksheet>
</Workbook>;
