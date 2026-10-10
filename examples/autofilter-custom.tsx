import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet
    name="Filtered inventory"
    autoFilter="A1:B4"
    autoFilterColumns={[
      { column: 0, includeBlanks: true },
      {
        column: 1,
        customFilters: {
          logic: "and",
          conditions: [
            { operator: "greaterThanOrEqual", value: "5" },
            { operator: "lessThan", value: "20" }
          ]
        }
      }
    ]}
  >
    <Row><Cell>Product</Cell><Cell>Quantity</Cell></Row>
    <Row><Cell>Apples</Cell><Cell value={12} /></Row>
    <Row><Cell>Oranges</Cell><Cell value={8} /></Row>
    <Row><Cell>Bananas</Cell><Cell value={3} /></Row>
  </Worksheet>
</Workbook>;
