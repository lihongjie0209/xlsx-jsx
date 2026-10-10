import { Workbook, Worksheet, Column, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Collapsed groups">
    <Column min={2} max={3} hidden={true} outlineLevel={1} />
    <Column min={4} max={4} collapsed={true} />
    <Row><Cell>Summary</Cell></Row>
    <Row hidden={true} outlineLevel={1}><Cell>Detail</Cell></Row>
    <Row collapsed={true}><Cell>Total</Cell></Row>
  </Worksheet>
</Workbook>;
