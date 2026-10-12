import { Workbook, Worksheet, Table, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Sales">
    <Table
      range="A1:B4"
      name="SalesTable"
      style="TableStyleMedium4"
      autoFilterColumns={[
        { column: 0, values: ["Tea", "Coffee"] },
        {
          column: 1,
          customFilters: {
            logic: "and",
            conditions: [
              { operator: "greaterThanOrEqual", value: "10" },
              { operator: "lessThan", value: "20" },
            ],
          },
        },
      ]}
    />
    <Row><Cell>Product</Cell><Cell>Amount</Cell></Row>
    <Row><Cell>Tea</Cell><Cell value={12} /></Row>
    <Row><Cell>Coffee</Cell><Cell value={18} /></Row>
    <Row><Cell>Juice</Cell><Cell value={9} /></Row>
  </Worksheet>
</Workbook>;
