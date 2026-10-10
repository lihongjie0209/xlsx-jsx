import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Rows">
    <Row height={24}>
      <Cell>Header</Cell>
    </Row>
    <Row />
    <Row hidden={true}>
      <Cell>Hidden row</Cell>
    </Row>
    <Row height={15.75} hidden={true} />
  </Worksheet>
</Workbook>;
