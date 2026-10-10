import { Workbook, Worksheet, Row, Cell } from "xlsx-jsx";

export default <Workbook>
  <Worksheet name="Font colors">
    <Row>
      <Cell fontColor="#C00000">RGB font color</Cell>
      <Cell value={42} fontColor="#4472C4" bold={true} fillColor="#DDEEFF" />
      <Cell formula="=1+1" fontColor="#548235" numberFormat="0.00" />
    </Row>
  </Worksheet>
</Workbook>;
