import { Workbook, Worksheet, Table, TableColumn, Row, Cell } from "xlsx-jsx";

export default (
  <Workbook>
    <Worksheet name="Products">
      <Table range="A1:B3" name="Products" style="TableStyleMedium2" altText="Product pricing data" altTextTitle="Product prices">
        <TableColumn />
        <TableColumn
          format={{
            strikethrough: true,
            underline: "double",
            script: "superscript",
            numberFormat: "#,##0.00",
            borderTop: "thin",
            borderTopColor: "#123456",
            borderBottom: "dashed",
            borderDiagonalType: "upDown",
          }}
        />
      </Table>
      <Row>
        <Cell>{"Product"}</Cell>
        <Cell>{"Price"}</Cell>
      </Row>
      <Row>
        <Cell>{"Tea"}</Cell>
        <Cell>{12}</Cell>
      </Row>
      <Row>
        <Cell>{"Coffee"}</Cell>
        <Cell>{18}</Cell>
      </Row>
    </Worksheet>
  </Workbook>
);
