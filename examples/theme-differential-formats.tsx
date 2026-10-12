import {
  Workbook,
  Worksheet,
  Row,
  Cell,
  Table,
  TableColumn,
  ConditionalFormat,
} from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Theme DXF">
    <ConditionalFormat range="A2:A5" type="cellIs" operator="greaterThan"
      formula1="0" fontColor={{ theme: 4, shade: 2 }}
      fillColor={{ theme: 5, shade: 1 }} borderTop="thin"
      borderTopColor={{ theme: 6, shade: 3 }} />
    <Table range="A3:A4" name="ThemeTable">
      <TableColumn format={{
        fontColor: { theme: 4, shade: 2 },
        fillColor: { theme: 5, shade: 1 },
        borderTop: "thin",
        borderTopColor: { theme: 6, shade: 3 },
      }} />
    </Table>
    <Row><Cell>Conditional formatting</Cell></Row>
    <Row><Cell value={15} /></Row>
    <Row><Cell>Table header</Cell></Row>
    <Row><Cell value={42} /></Row>
    <Row><Cell value={43} /></Row>
  </Worksheet>
</Workbook>;
