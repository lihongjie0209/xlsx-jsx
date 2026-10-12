import { Workbook, Worksheet, Table, TableColumn, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Inventory">
      <Table range="A1:B3" name="InventoryTable">
        <TableColumn format={{ horizontalAlign: "left", verticalAlign: "center", wrapText: true }} />
        <TableColumn format={{ horizontalAlign: "right", verticalAlign: "top" }} />
      </Table>
      <Row><Cell>{"Description"}</Cell><Cell>{"Quantity"}</Cell></Row>
      <Row><Cell>{"A long product description"}</Cell><Cell>{12}</Cell></Row>
      <Row><Cell>{"Another item"}</Cell><Cell>{8}</Cell></Row>
    </Worksheet>
  </Workbook>
);
